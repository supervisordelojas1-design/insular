# Insular

App web (PWA) para acompanhar glicemia e aplicações de insulina na gestação.
Funciona no celular como aplicativo (Adicionar à Tela de Início), salva os registros
no aparelho e sincroniza automaticamente com uma planilha Google via Apps Script
(`apps-script/Code.gs`).

- `index.html` — o app
- `sw.js` — cache para uso sem internet (mude `CACHE` ao atualizar o app)
- `manifest.webmanifest` + ícones — instalação como app
- `apps-script/Code.gs` — script da planilha (publicado como app da web; troque o TOKEN)

A URL do script e a senha são informadas em Ajustes, dentro do app.
