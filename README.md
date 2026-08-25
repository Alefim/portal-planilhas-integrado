# Portal de Planilhas — integração direta com Google Drive

Projeto para Google Apps Script conectado diretamente à planilha nativa do Google Sheets no Google Drive.

## Planilha principal do Drive

- Arquivo: `PLAN_COMITE_Sergio_Aguiar_2026 Ok`
- ID do Google Drive/Sheets: `1jabSU0G5IYoKP6raGtZszQuV8Ti92MJu9E44zkOWrCc`
- Aba inicial: `RESUMO GERAL`
- gid: `349757095`

O portal não usa arquivo `.xlsx` local. O `Code.gs` executa `SpreadsheetApp.openById(...)` para validar e resolver a planilha diretamente no Google Drive da conta que executa o Apps Script.

## Controle e auditoria

Ao executar `configurarNovaConta()` uma única vez, o Apps Script cria no Google Drive da conta executora:

- `Portal de Planilhas — Controle de Acessos`
- `Portal de Planilhas — Auditoria de Acessos`

Também configura usuários, sessões, logins, acessos e a planilha principal.

Além disso, cada sessão é espelhada automaticamente na planilha consolidada:

- `Registro de Sessões - Portal de Planilhas`
- ID: `1Q4QHcoPY27fboJEfq73pYgRrntG50N6c8rAx2VdqtHw`
- Aba: `Registros`

São gravados: usuário, entrada, saída, duração, número de aberturas, status, dispositivo e origem.

## Publicação

1. Abra/crie o projeto do Google Apps Script estando na conta Google que terá acesso à planilha principal.
2. Copie `Code.gs`, `Index.html` e `appsscript.json`.
3. Execute `configurarNovaConta()` uma vez e autorize o Google.
4. Execute `diagnosticoSistema()` para conferir.
5. Implante como Aplicativo da Web, executando como a conta proprietária/autorizada.
6. Após atualizar os usuários, execute `migrarCredenciais2026()` e publique uma nova versão da implantação.

> O endereço do GitHub Pages oferece somente a interface estática de contingência. Para registrar logins, duração, heartbeat, abertura de planilhas e encerramento na planilha de sessões, use o endereço da implantação do Google Apps Script.

Se a conta não tiver permissão para o arquivo principal, o sistema informa que a planilha do Google Drive não está acessível.

## Botão de acesso direto

O botão **Abrir planilha →** abre o arquivo original do Google Drive:

`https://docs.google.com/spreadsheets/d/1jabSU0G5IYoKP6raGtZszQuV8Ti92MJu9E44zkOWrCc/edit?usp=drivesdk`

Se o usuário já estiver autenticado no Google, a planilha abre diretamente. Caso contrário, o próprio Google solicita a conta e, após o login, abre o arquivo conforme as permissões dessa conta.
