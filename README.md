## 🚀 Funcionalidades

- Filtro por:
  - Data de emissão
  - Produto (intervalo)
- Validação de permissão de usuário via parâmetro (MV)
- Consulta otimizada via SQL (SC7 e SB1)
- Geração automática de planilha Excel
- Exibição de progresso durante o processamento
- Conversão de moeda (código → descrição)

## 📊 Dados apresentados

O relatório contempla as seguintes informações:

- Número do pedido
- Item
- Data de emissão
- Fornecedor
- Produto
- Part Number (fabricante)
- Descrição
- Unidade de medida
- Quantidade
- Valor unitário
- IPI
- Valor total
- Data de entrega
- Centro de custo
- Solicitação de compra (SC)
- Moeda e taxa de conversão

## 🛠️ Tecnologias utilizadas

- ADVPL (TOTVS Protheus)
- FWMSExcel para geração de Excel
- SQL embarcado (TopConn)

## 🔐 Controle de acesso

O acesso ao relatório é controlado via parâmetro:
