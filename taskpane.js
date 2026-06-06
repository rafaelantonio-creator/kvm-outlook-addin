Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    aplicarTemplateSeNecessario();
    var btn = document.getElementById('btnAplicar');
    if (btn) {
      btn.addEventListener('click', function() { inserirTemplate(); });
    }
  }
});

function aplicarTemplateSeNecessario() {
  var item = Office.context.mailbox.item;
  if (!item) return;

  item.body.getAsync(Office.CoercionType.Html, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Só insere se o template ainda não foi aplicado
      if (result.value.indexOf('Convite Assertivo') === -1) {
        inserirTemplate();
      }
    }
  });
}

function inserirTemplate() {
  var template = [
    '<html><body style="font-family: Segoe UI, sans-serif; font-size: 14px; color: #333;">',
    '<p><strong style="color:#0078D4;">🎯 Objetivo da reunião:</strong><br>',
    '[Descreva claramente o objetivo principal desta reunião]</p>',
    '<p><strong style="color:#0078D4;">✅ Resultado esperado:</strong><br>',
    '[O que esperamos decidir, aprovar ou entregar ao final desta reunião]</p>',
    '<p><strong style="color:#0078D4;">👥 Participantes e papéis:</strong><br>',
    '[Nome] — [Papel/responsabilidade]<br>',
    '[Nome] — [Papel/responsabilidade]</p>',
    '<p><strong style="color:#0078D4;">📋 Pauta:</strong><br>',
    '1. [Tópico 1]<br>',
    '2. [Tópico 2]<br>',
    '3. [Tópico 3]</p>',
    '<p><strong style="color:#0078D4;">⏱ Duração estimada por tópico:</strong><br>',
    '[Ex: 10min abertura | 20min discussão | 10min próximos passos]</p>',
    '<p><strong style="color:#0078D4;">📎 Materiais de apoio:</strong><br>',
    '[Links ou anexos relevantes para a reunião]</p>',
    '<hr style="border: none; border-top: 1px solid #e0e0e0; margin: 16px 0;">',
    '<p style="font-size: 11px; color: #999;">',
    'Template gerado automaticamente pelo Add-in Convite Assertivo — KVM Tecnologia',
    '</p>',
    '</body></html>'
  ].join('');

  Office.context.mailbox.item.body.setAsync(
    template,
    { coercionType: Office.CoercionType.Html },
    function(result) {
      var status = document.getElementById('status');
      if (!status) return;
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        status.textContent = '✅ Template aplicado!';
        status.style.color = 'green';
        status.style.display = 'block';
      } else {
        status.textContent = '❌ Erro: ' + result.error.message;
        status.style.color = 'red';
        status.style.display = 'block';
      }
    }
  );
}
