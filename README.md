# Ar Condicionado · Home Center Romanos

Controle de estoque de ar-condicionado por loja, com baixa, entrada e estorno.

- **Site:** https://ismael9zxx-maker.github.io/ar-condicionado/ (GitHub Pages, ramo `main`)
- **Dados:** planilha "Frio Estoque 2026" no Google Drive (abas `Estoque` e `Lançamentos`)
- **Ponte site ↔ planilha:** Apps Script em `apps-script/Codigo.gs`, publicado como App da Web

O link `/exec` do App da Web fica na constante `API_URL`, no começo do script em `index.html`.

Na aba `Lançamentos`, a coluna K guarda o motivo da baixa (Venda, Avaria ou Transferência). Só as baixas por venda entram na aba Vendas. Ao mudar `apps-script/Codigo.gs`, cole o arquivo de novo no Apps Script e publique uma nova versão da implantação.
