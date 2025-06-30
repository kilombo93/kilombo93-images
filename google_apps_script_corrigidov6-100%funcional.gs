function doPost(e) {
  const SPREADSHEET_ID = '1kDBXT-LHpmhikhj6RmOECeemb6qz2FgkWhmZzeT6hjI'; // Substitua pelo ID da sua planilha
  
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    // Cabeçalhos da planilha (conforme especificação)
    const headers = [
      'Timestamp', 'Código do Pedido', 'Nome da Empresa', 'CNPJ', 'Nome do Responsável',
      'E-mail', 'Telefone', 'Cidade', 'Estado', 'Qtd de Lojas', 'Vendedores/Loja',
      'Site da Loja', 'Instagram', 'TikTok', 'Forma Pagamento', 'Nome do Produto',
      'SKU', 'Tamanho', 'Quantidade', 'Preço Unitário', 'Subtotal Produto', 'Total do Pedido',
      'Lucro Estimado', 'Benefícios Escolhidos', 'Observações'
    ];

    // Se a planilha estiver vazia, adicione os cabeçalhos
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      
      // Formatação dos cabeçalhos
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4CAF50');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }

    // Extrair totais globais do pedido do primeiro item (já vêm do HTML)
    const firstItem = data[0];
    const orderCode = firstItem.codigoPedido; 
    const totalPedidoGlobal = firstItem.totalPedido; 

    // Processar cada produto como uma linha separada
    data.forEach(item => {
      const rowData = [
        new Date(), // Timestamp (usar o do servidor para consistência)
        item.codigoPedido, // Código do Pedido (já vem do HTML)
        item.nomeEmpresa || '', // Nome da Empresa
        item.cnpj || '', // CNPJ
        item.nomeResponsavel || '', // Nome do Responsável
        item.email || '', // E-mail
        item.telefone || '', // Telefone
        item.cidade || '', // Cidade
        item.estado || '', // Estado
        item.quantidadeLojas || '', // Qtd de Lojas
        item.vendedoresPorLoja || '', // Vendedores/Loja
        item.website || '', // Site da Loja
        item.instagram || '', // Instagram
        item.tiktok || '', // TikTok
        item.formaPagamento || '', // Forma Pagamento
        item.nomeProduto || '', // Nome do Produto
        item.sku || '', // SKU
        item.tamanho || '', // Tamanho (novo campo)
        item.quantidadeSolicitada, // Quantidade (já é número)
        item.precoUnitario, // Preço Unitário (já é número)
        item.totalProduto, // Subtotal Produto (já é número)
        totalPedidoGlobal, // Total do Pedido (usar o global para todas as linhas)
        item.lucroEstimado, // Lucro Estimado (já é número)
        item.beneficiosEscolhidos || '', // Benefícios Escolhidos
        item.observacoes || '' // Observações
      ];

      // Inserir nova linha no topo (linha 2, logo após os cabeçalhos)
      sheet.insertRowBefore(2);
      sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    });

    // Aplicar formatação às colunas de valores monetários
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Formatar colunas de valores (colunas 20, 21, 22, 23 - Preço Unitário, Subtotal Produto, Total do Pedido, Lucro Estimado)
      // As posições das colunas mudaram, então precisamos reajustar
      const moneyColumns = [20, 21, 22, 23]; // Preço Unitário, Subtotal Produto, Total do Pedido, Lucro Estimado
      moneyColumns.forEach(col => {
        const range = sheet.getRange(2, col, lastRow - 1, 1);
        range.setNumberFormat('R$ #,##0.00');
      });
    }

    return ContentService.createTextOutput(JSON.stringify({
      'status': 'success',
      'message': 'Dados recebidos com sucesso!',
      'orderCode': orderCode
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Erro no processamento:', error);
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': 'Erro ao processar dados: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Este script é para receber dados via POST.').setMimeType(ContentService.MimeType.TEXT);
}

// Função auxiliar para testar a conexão
function testConnection() {
  const SPREADSHEET_ID = '1kDBXT-LHpmhikhj6RmOECeemb6qz2FgkWhmZzeT6hjI';
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    console.log('Conexão com a planilha estabelecida com sucesso!');
    console.log('Nome da planilha:', sheet.getName());
    return true;
  } catch (error) {
    console.error('Erro ao conectar com a planilha:', error);
    return false;
  }
}

