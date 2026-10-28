# Quebra FLV

App de celular para lançar quebras de FLV (hortifruti) da loja e imprimir no layout da **Planilha Padrão de Lançamentos** da Prevenção de Perdas.

- Busca por nome ou código interno (947 produtos das estruturas Alho/Frutas/Legumes, Folhagens/Orgânicos, Manipulados e Ovos)
- Leitura de código de barras pela câmera (etiqueta da balança e EAN do fornecedor), com vínculo ao produto na primeira leitura
- Soma de pesos com "+"
- Cadastro de produtos que não estão na relação
- Impressão / PDF / compartilhar, com tipo (Avarias, Reaproveitamento, Refeitório) e data
- Funciona sem internet depois da primeira abertura; os dados ficam salvos no próprio celular (com backup em arquivo)

Publicado com GitHub Pages: Settings → Pages → Branch `main` / pasta `/ (root)`.

Para publicar uma versão nova, altere `VERSAO` em `sw.js`.

## Aproveitamento do refeitório (envio para a planilha do Drive)

Itens marcados com **R** no app vão para o *Controle de Aproveitamento Diário - FLV*.
O script `apps-script/Aproveitamento.gs` fica dentro da planilha do Google e grava na aba do dia
(código, descrição, quantidade e fórmulas de custo/valor pela aba CUSTO).

Instalação (uma vez):
1. Abra a planilha no Google Planilhas → **Extensões → Apps Script**.
2. Apague o que estiver no editor, cole o conteúdo de `Aproveitamento.gs` e salve.
3. **Implantar → Nova implantação** → tipo **App da Web** → Executar como **Eu** → Quem pode acessar **Qualquer pessoa** → Implantar → autorize.
4. Copie o endereço que termina em `/exec` e cole no app em **Enviar aproveitamento → Ligar à planilha**.
