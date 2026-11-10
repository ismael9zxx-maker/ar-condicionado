# Ar Condicionado · Home Center Romanos

Controle de estoque de ar-condicionado por loja, com baixa, entrada e estorno.

- **Site:** https://ismael9zxx-maker.github.io/ar-condicionado/ (GitHub Pages, ramo `main`)
- **Dados:** planilha "Frio Estoque 2026" no Google Drive (abas `Estoque` e `Lançamentos`)
- **Ponte site ↔ planilha:** Apps Script em `apps-script/Codigo.gs`, publicado como App da Web

O link `/exec` do App da Web fica na constante `API_URL`, no começo do script em `index.html`.

Na aba `Lançamentos`, a coluna K guarda o motivo da baixa (Venda, Avaria ou Transferência). Só as baixas por venda entram na aba Vendas. Ao mudar `apps-script/Codigo.gs`, cole o arquivo de novo no Apps Script e publique uma nova versão da implantação.

A aba `Histórico` (criada sozinha pelo script) recebe as vendas de meses anteriores ao app, uma linha por mês, loja e produto: `Mês (AAAA-MM)`, `Loja`, `Código`, `Quantidade`. A aba Vendas soma isso com as baixas por venda do app. Modelos novos são cadastrados pela área da gerência no site.
