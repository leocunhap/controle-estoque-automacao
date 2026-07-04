/**
 * AGENTE DE ALERTA DE ESTOQUE → E-MAIL (GMAIL, formatado em HTML)
 * --------------------------------------------------------------------
 * Como instalar:
 * 1. Na planilha do Google Sheets, vá em Extensões > Apps Script.
 * 2. Cole este código (substitua o conteúdo do arquivo Code.gs).
 * 3. Substitua DESTINATARIOS pelos e-mails que devem receber o alerta.
 * 4. (Opcional, para IA) Coloque sua API Key da Anthropic em
 *    Propriedades do Projeto > Editor de propriedades do script,
 *    como Script Property chamada ANTHROPIC_API_KEY.
 * 5. Rode a função "instalarGatilhos" UMA VEZ manualmente. Isso
 *    configura os dois gatilhos e pede as permissões necessárias.
 *
 * OS DOIS GATILHOS CONFIGURADOS:
 * a) TEMPO REAL: toda vez que a aba "Estoque" for editada, o script
 *    aguarda 15 segundos (debounce) e então verifica se a lista de
 *    itens críticos mudou. Só envia e-mail se mudou.
 * b) HORÁRIO FIXO: roda 1x por dia às 8h como verificação de segurança
 *    (cobre edições via API/importação que não disparam o onEdit).
 *
 * LÓGICA DE ENVIO:
 * ✅ Aguarda 15s após a última edição antes de verificar (debounce).
 * ✅ Envia somente se a lista de itens críticos mudou.
 * ✅ Não reenvia o mesmo alerta se nada mudou.
 * ✅ Se um novo produto entrar em crítico, envia imediatamente.
 * ✅ Se todos saírem do crítico, reseta o estado — próxima entrada
 *    dispara novo e-mail normalmente.
 *
 * VISUAL: tema escuro (navy) alinhado ao dashboard "Controle de Estoque"
 * (Looker Studio) — fundo #0f1729, cards com borda sutil e badges
 * coloridos de urgência.
 *
 * Limite gratuito do Gmail: até 100 e-mails/dia (@gmail.com comum).
 */

const DESTINATARIOS  = "seuemail@gmail.com,outroemail@gmail.com"; // <- substitua pelos e-mails que devem receber o alerta
const NOME_ABA_ESTOQUE  = "Estoque";
const USAR_IA_PARA_RESUMO = true;
const DEBOUNCE_SEGUNDOS  = 15; // tempo de espera após a última edição antes de verificar

// ─────────────────────────────────────────────────────────────────────
// INSTALAÇÃO DOS GATILHOS
// Execute esta função UMA VEZ manualmente para configurar tudo.
// Não precisa rodar de novo, a menos que queira recriar os gatilhos.
// ─────────────────────────────────────────────────────────────────────
function instalarGatilhos() {
  // Remove gatilhos antigos deste projeto para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Gatilho 1: tempo real, ao editar a planilha
  ScriptApp.newTrigger("aoEditarPlanilha")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Gatilho 2: horário fixo, 1x por dia às 8h
  ScriptApp.newTrigger("verificarEstoqueEAlertar")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("Gatilhos instalados: onEdit (tempo real) + diário às 8h.");
}

// ─────────────────────────────────────────────────────────────────────
// DEBOUNCE
// Cada edição cancela o timer anterior e cria um novo de 15 segundos.
// O e-mail só é disparado quando o usuário para de editar.
// Isso evita múltiplos envios durante uma sessão de edição.
// ─────────────────────────────────────────────────────────────────────
function aoEditarPlanilha(e) {
  const aba = e.range.getSheet();
  if (aba.getName() !== NOME_ABA_ESTOQUE) return;

  const props = PropertiesService.getScriptProperties();

  // Cancela o trigger de debounce anterior, se existir
  const triggerIdAnterior = props.getProperty("DEBOUNCE_TRIGGER_ID");
  if (triggerIdAnterior) {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getUniqueId() === triggerIdAnterior) ScriptApp.deleteTrigger(t);
    });
  }

  // Cria um novo trigger para rodar daqui a DEBOUNCE_SEGUNDOS
  const novoTrigger = ScriptApp.newTrigger("executarAposDebounce")
    .timeBased()
    .after(DEBOUNCE_SEGUNDOS * 1000)
    .create();

  props.setProperty("DEBOUNCE_TRIGGER_ID", novoTrigger.getUniqueId());
}

// Chamada pelo trigger de debounce após o tempo de espera.
// Faz a limpeza do trigger antes de executar a verificação.
function executarAposDebounce() {
  const props = PropertiesService.getScriptProperties();

  // Auto-limpeza: remove o trigger que chamou esta função
  const triggerId = props.getProperty("DEBOUNCE_TRIGGER_ID");
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
  });
  props.deleteProperty("DEBOUNCE_TRIGGER_ID");

  verificarEstoqueEAlertar();
}

// ─────────────────────────────────────────────────────────────────────
// FINGERPRINT
// Gera uma string com os nomes dos itens críticos em ordem alfabética.
// Compara com a última salva: só envia e-mail se a lista mudou.
// Se todos os itens saírem do crítico, o fingerprint é apagado —
// garantindo que uma nova entrada em crítico dispare normalmente.
// ─────────────────────────────────────────────────────────────────────
function gerarFingerprint(itensCriticos) {
  return itensCriticos
    .map(i => i.nome)
    .sort()
    .join("|");
}

function verificarEstoqueEAlertar() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(NOME_ABA_ESTOQUE);
  const dados   = aba.getDataRange().getValues();
  const headers = dados[0];

  const idxNome      = headers.indexOf("Nome_Produto");
  const idxCategoria = headers.indexOf("Categoria");
  const idxQtd       = headers.indexOf("Quantidade_Atual");
  const idxMin       = headers.indexOf("Estoque_Minimo");
  const idxStatus    = headers.indexOf("Status");
  const idxFornecedor = headers.indexOf("Fornecedor");

  const itensCriticos = [];
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (linha[idxStatus] === "Crítico") {
      const atual  = Number(linha[idxQtd]);
      const minimo = Number(linha[idxMin]);
      itensCriticos.push({
        nome:            linha[idxNome],
        categoria:       linha[idxCategoria],
        atual,
        minimo,
        fornecedor:      linha[idxFornecedor],
        percentualFalta: minimo > 0 ? ((minimo - atual) / minimo) * 100 : 0
      });
    }
  }

  const props = PropertiesService.getScriptProperties();

  if (itensCriticos.length === 0) {
    Logger.log("Nenhum item crítico. Nenhum alerta enviado.");
    // Reseta o fingerprint para que o próximo item crítico dispare novo e-mail
    props.deleteProperty("FINGERPRINT_CRITICOS");
    return;
  }

  // Compara fingerprint atual com o salvo
  const fingerprintAtual = gerarFingerprint(itensCriticos);
  const fingerprintSalvo = props.getProperty("FINGERPRINT_CRITICOS");

  if (fingerprintAtual === fingerprintSalvo) {
    Logger.log("Lista de críticos não mudou. Nenhum alerta enviado.");
    return;
  }

  // Lista mudou → salva novo fingerprint e envia o e-mail
  props.setProperty("FINGERPRINT_CRITICOS", fingerprintAtual);

  itensCriticos.sort((a, b) => b.percentualFalta - a.percentualFalta);

  let resumoIA = null;
  if (USAR_IA_PARA_RESUMO) resumoIA = gerarResumoComIA(itensCriticos);

  const htmlEmail = montarEmailHtml(itensCriticos, resumoIA);
  enviarEmail(htmlEmail, itensCriticos.length);
}

// ─────────────────────────────────────────────────────────────────────
// URGÊNCIA
// Define cor e etiqueta com base em quanto o estoque está abaixo
// do mínimo. Usado nos badges coloridos da tabela do e-mail.
// Cores calibradas para fundo escuro (navy).
// ─────────────────────────────────────────────────────────────────────
function obterUrgencia(percentualFalta) {
  if (percentualFalta >= 70) return { label: "URGENTE",  cor: "#f87171", fundo: "#3a1a1f" };
  if (percentualFalta >= 40) return { label: "ALTO",     cor: "#fb923c", fundo: "#3a2410" };
  return                            { label: "MODERADO", cor: "#fbbf24", fundo: "#332a0c" };
}

// ─────────────────────────────────────────────────────────────────────
// RESUMO COM IA (OPCIONAL)
// Usa a API da Anthropic para gerar um parágrafo executivo no topo
// do e-mail. Se a chave não estiver configurada ou a chamada falhar,
// o e-mail segue normalmente sem esse bloco.
// ─────────────────────────────────────────────────────────────────────
function gerarResumoComIA(itens) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    Logger.log("ANTHROPIC_API_KEY não configurada. E-mail seguirá sem resumo executivo.");
    return null;
  }

  const listaItens = itens.map(i =>
    `- ${i.nome} | Categoria: ${i.categoria} | Atual: ${i.atual} | Mínimo: ${i.minimo} | Fornecedor: ${i.fornecedor}`
  ).join("\n");

  const prompt = `Você é um agente de gestão de estoque. Abaixo está a lista de produtos em estado crítico, ordenada do mais urgente para o menos urgente. Escreva um resumo executivo curto (máximo 4 frases, texto simples, sem markdown) destacando os 2-3 itens mais urgentes e sugerindo a ação de reposição prioritária.\n\nLista:\n${listaItens}`;

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 300,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
    const json = JSON.parse(response.getContentText());
    return json.content?.[0]?.text ?? null;
  } catch (err) {
    Logger.log("Erro ao chamar a IA: " + err.message);
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────
// MONTAGEM DO E-MAIL EM HTML — TEMA ESCURO (alinhado ao dashboard)
// Cabeçalho com data, KPIs (total críticos / urgentes), bloco opcional
// do resumo da IA e tabela com badge de urgência por produto.
// ─────────────────────────────────────────────────────────────────────
function montarEmailHtml(itens, resumoIA) {
  const dataFormatada = new Date().toLocaleDateString("pt-BR", {
    weekday: "long", year: "numeric", month: "long", day: "numeric"
  });
  const urgentes = itens.filter(i => i.percentualFalta >= 70).length;

  let linhasTabela = "";
  itens.forEach((item, idx) => {
    const urgencia = obterUrgencia(item.percentualFalta);
    const corLinha = idx % 2 === 0 ? "#16213a" : "#131c31";
    linhasTabela += `
      <tr style="background-color:${corLinha};">
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;white-space:nowrap;">
          <span style="display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;color:${urgencia.cor};background-color:${urgencia.fundo};">
            ${urgencia.label}
          </span>
        </td>
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;font-weight:600;color:#e8ecf5;white-space:nowrap;">${item.nome}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;color:#8b9bc1;white-space:nowrap;">${item.categoria}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;text-align:center;color:#ef4444;font-weight:bold;white-space:nowrap;">${item.atual}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;text-align:center;color:#8b9bc1;white-space:nowrap;">${item.minimo}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #232f4d;color:#8b9bc1;white-space:nowrap;">${item.fornecedor}</td>
      </tr>`;
  });

  const blocoResumoIA = resumoIA ? `
    <div style="background-color:#152238;border-left:4px solid #4a7fe8;border-radius:8px;padding:14px 18px;margin:0 0 24px 0;">
      <p style="margin:0;color:#c7d6f5;font-size:14px;line-height:1.5;">
        <strong style="color:#4a7fe8;">📋 Resumo:</strong> ${resumoIA}
      </p>
    </div>` : "";

  return `<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<style>
  body { margin:0; padding:0; background-color:#0f1729; }
  table { border-collapse: collapse; }
  @media only screen and (max-width: 600px) {
    .email-outer-pad { padding-left:12px !important; padding-right:12px !important; }
    .email-header-pad { padding:18px 18px !important; }
    .email-body-pad { padding:16px 18px 20px !important; }
    .header-title { font-size:19px !important; }
    .kpi-row { display:block !important; }
    .kpi-cell { display:block !important; width:100% !important; box-sizing:border-box; padding:5px 0 !important; }
    .kpi-value { font-size:26px !important; }
    .footer-pad { padding:14px 18px !important; }
  }
</style>
</head>
<body style="margin:0;padding:0;background-color:#0f1729;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#0f1729;">
    <tr>
      <td align="center" style="padding:16px 8px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;background-color:#0f1729;border:1px solid #232f4d;border-radius:12px;">
          <tr>
            <td style="padding:0;">

              <!-- HEADER -->
              <div class="email-header-pad" style="background-color:#131c31;padding:24px 28px;border-bottom:1px solid #232f4d;border-radius:12px 12px 0 0;">
                <p class="header-title" style="margin:0;color:#e8ecf5;font-size:22px;font-weight:bold;">
                  <span style="color:#4a7fe8;">🤖</span> AGENTE DE ESTOQUE
                </p>
                <p style="margin:4px 0 0 0;color:#8b9bc1;font-size:13px;font-style:italic;">Alerta de itens críticos · ${dataFormatada}</p>
              </div>

              <!-- KPI CARDS -->
              <div class="email-outer-pad" style="padding:22px 28px 6px;">
                <table class="kpi-row" width="100%" cellpadding="0" cellspacing="0">
                  <tr class="kpi-row">
                    <td class="kpi-cell" width="50%" style="padding:6px;">
                      <div style="background-color:#1c1418;border:1px solid #4a2224;border-radius:10px;padding:16px;text-align:center;">
                        <p style="margin:0 0 4px;font-size:11px;font-weight:700;color:#f19999;text-transform:uppercase;letter-spacing:0.06em;">Itens Críticos</p>
                        <p class="kpi-value" style="margin:0;font-size:30px;font-weight:800;color:#ef4444;">${itens.length}</p>
                      </div>
                    </td>
                    <td class="kpi-cell" width="50%" style="padding:6px;">
                      <div style="background-color:#231a0c;border:1px solid #543f16;border-radius:10px;padding:16px;text-align:center;">
                        <p style="margin:0 0 4px;font-size:11px;font-weight:700;color:#f0c98a;text-transform:uppercase;letter-spacing:0.06em;">Urgentes (≥70%)</p>
                        <p class="kpi-value" style="margin:0;font-size:30px;font-weight:800;color:#fbbf24;">${urgentes}</p>
                      </div>
                    </td>
                  </tr>
                </table>
              </div>

              <!-- BODY -->
              <div class="email-body-pad email-outer-pad" style="padding:16px 28px 26px;">
                ${blocoResumoIA}
                <p style="margin:0 0 12px;font-size:11px;font-weight:700;color:#4a7fe8;text-transform:uppercase;letter-spacing:0.10em;">Itens Abaixo do Estoque Mínimo</p>
                <div style="border:1px solid #232f4d;border-radius:10px;overflow-x:auto;-webkit-overflow-scrolling:touch;">
                  <table style="width:100%;min-width:560px;border-collapse:collapse;font-size:13px;">
                    <thead>
                      <tr style="background-color:#1e2d4d;">
                        <th style="padding:10px 12px;text-align:left;color:#a8c0ec;font-size:12px;white-space:nowrap;">Urgência</th>
                        <th style="padding:10px 12px;text-align:left;color:#a8c0ec;font-size:12px;white-space:nowrap;">Produto</th>
                        <th style="padding:10px 12px;text-align:left;color:#a8c0ec;font-size:12px;white-space:nowrap;">Categoria</th>
                        <th style="padding:10px 12px;text-align:center;color:#a8c0ec;font-size:12px;white-space:nowrap;">Atual</th>
                        <th style="padding:10px 12px;text-align:center;color:#a8c0ec;font-size:12px;white-space:nowrap;">Mínimo</th>
                        <th style="padding:10px 12px;text-align:left;color:#a8c0ec;font-size:12px;white-space:nowrap;">Fornecedor</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${linhasTabela}
                    </tbody>
                  </table>
                </div>
                <p style="margin:8px 2px 0;font-size:11px;color:#5c6d95;">↔ Arraste a tabela para os lados para ver todas as colunas.</p>
              </div>

              <!-- FOOTER -->
              <div class="footer-pad" style="padding:16px 28px;background-color:#131c31;border-top:1px solid #232f4d;text-align:center;border-radius:0 0 12px 12px;">
                <p style="margin:0;color:#5c6d95;font-size:12px;">
                  Gerado automaticamente pelo Agente de Estoque · Confira o dashboard completo no Looker Studio
                </p>
              </div>

            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

// ─────────────────────────────────────────────────────────────────────
// ENVIO DO E-MAIL
// MailApp é nativo do Google — não precisa de senha, token ou serviço
// externo. Inclui fallback em texto puro para clientes sem suporte HTML.
// ─────────────────────────────────────────────────────────────────────
function enviarEmail(corpoHtml, totalCriticos) {
  const assunto = `🤖 Agente de Estoque — ${totalCriticos} item(ns) crítico(s) — ${new Date().toLocaleDateString("pt-BR")}`;
  MailApp.sendEmail({
    to: DESTINATARIOS,
    subject: assunto,
    htmlBody: corpoHtml,
    body: "Seu cliente de e-mail não exibe HTML. Acesse a planilha para ver os itens críticos."
  });
  Logger.log("E-mail enviado para: " + DESTINATARIOS);
}
