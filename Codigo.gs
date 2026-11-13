// ============================================================
// CÓDIGO.GS - Backend do Aplicativo de Biblioteca
// ============================================================

const SPREADSHEET_ID = '1qdYUOgmapLJUtuAsEY6TO8deNaVq204fw2F-KouLevE';
// Exemplo: '1a2b3c4d5e6f7g8h9i0j...'

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Biblioteca Web App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- HELPERS ----------

function getSheet_(nomeAba) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nomeAba);
}

function gerarId_(prefixo) {
  return prefixo + '-' + new Date().getTime().toString(36).toUpperCase();
}

// ---------- CATALOGAÇÃO ----------

function cadastrarLivro(dados) {
  try {
    const sheet = getSheet_('Livros');
    const id = gerarId_('LIV');
    const agora = new Date().toLocaleString('pt-BR');
    
    sheet.appendRow([
      id,
      dados.titulo,
      dados.autor,
      dados.isbn || '',
      dados.ano || '',
      'Disponível',
      dados.localizacao || '',
      agora
    ]);
    
    return { sucesso: true, mensagem: 'Livro cadastrado! ID: ' + id, id: id };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

function buscarLivros(termo) {
  const sheet = getSheet_('Livros');
  const dados = sheet.getDataRange().getValues();
  const resultado = [];
  
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const texto = (linha[1] + ' ' + linha[2] + ' ' + linha[3]).toLowerCase();
    if (!termo || texto.includes(termo.toLowerCase())) {
      resultado.push({
        id: linha[0],
        titulo: linha[1],
        autor: linha[2],
        isbn: linha[3],
        ano: linha[4],
        status: linha[5],
        localizacao: linha[6]
      });
    }
  }
  return resultado;
}

function listarLivrosDisponiveis() {
  return buscarLivros('').filter(l => l.status === 'Disponível');
}

// ---------- RETIRADA ----------

function realizarRetirada(idLivro, nomeLeitor, diasPrazo) {
  try {
    // 1. Verificar se livro existe e está disponível
    const sheetLivros = getSheet_('Livros');
    const dadosLivros = sheetLivros.getDataRange().getValues();
    let linhaLivro = -1;
    let tituloLivro = '';
    
    for (let i = 1; i < dadosLivros.length; i++) {
      if (dadosLivros[i][0] === idLivro) {
        linhaLivro = i + 1;
        tituloLivro = dadosLivros[i][1];
        if (dadosLivros[i][5] !== 'Disponível') {
          return { sucesso: false, mensagem: 'Livro não está disponível. Status: ' + dadosLivros[i][5] };
        }
        break;
      }
    }
    
    if (linhaLivro === -1) {
      return { sucesso: false, mensagem: 'Livro não encontrado.' };
    }
    
    // 2. Registrar empréstimo
    const sheetEmp = getSheet_('Emprestimos');
    const idEmp = gerarId_('EMP');
    const hoje = new Date();
    const prazo = new Date();
    prazo.setDate(prazo.getDate() + (parseInt(diasPrazo) || 7));
    
    sheetEmp.appendRow([
      idEmp,
      idLivro,
      tituloLivro,
      nomeLeitor,
      hoje.toLocaleDateString('pt-BR'),
      prazo.toLocaleDateString('pt-BR'),
      '',
      'Ativo'
    ]);
    
    // 3. Atualizar status do livro
    sheetLivros.getRange(linhaLivro, 6).setValue('Emprestado');
    
    return { sucesso: true, mensagem: 'Retirada registrada! Prazo: ' + prazo.toLocaleDateString('pt-BR'), id: idEmp };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// ---------- DEVOLUÇÃO ----------

function listarEmprestimosAtivos() {
  const sheet = getSheet_('Emprestimos');
  const dados = sheet.getDataRange().getValues();
  const ativos = [];
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][7] === 'Ativo') {
      ativos.push({
        id: dados[i][0],
        idLivro: dados[i][1],
        titulo: dados[i][2],
        leitor: dados[i][3],
        dataRetirada: dados[i][4],
        prazo: dados[i][5]
      });
    }
  }
  return ativos;
}

function realizarDevolucao(idEmprestimo) {
  try {
    const sheetEmp = getSheet_('Emprestimos');
    const dadosEmp = sheetEmp.getDataRange().getValues();
    let linhaEmp = -1;
    let idLivro = '';
    
    for (let i = 1; i < dadosEmp.length; i++) {
      if (dadosEmp[i][0] === idEmprestimo) {
        linhaEmp = i + 1;
        idLivro = dadosEmp[i][1];
        if (dadosEmp[i][7] !== 'Ativo') {
          return { sucesso: false, mensagem: 'Este empréstimo já foi finalizado.' };
        }
        break;
      }
    }
    
    if (linhaEmp === -1) {
      return { sucesso: false, mensagem: 'Empréstimo não encontrado.' };
    }
    
    // Atualizar empréstimo
    sheetEmp.getRange(linhaEmp, 7).setValue(new Date().toLocaleDateString('pt-BR'));
    sheetEmp.getRange(linhaEmp, 8).setValue('Devolvido');
    
    // Atualizar status do livro
    const sheetLivros = getSheet_('Livros');
    const dadosLivros = sheetLivros.getDataRange().getValues();
    for (let i = 1; i < dadosLivros.length; i++) {
      if (dadosLivros[i][0] === idLivro) {
        sheetLivros.getRange(i + 1, 6).setValue('Disponível');
        break;
      }
    }
    
    return { sucesso: true, mensagem: 'Devolução registrada com sucesso!' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// ---------- DAR BAIXA ----------

function darBaixa(idLivro, motivo, bibliotecario) {
  try {
    const sheetLivros = getSheet_('Livros');
    const dadosLivros = sheetLivros.getDataRange().getValues();
    let linhaLivro = -1;
    let tituloLivro = '';
    
    for (let i = 1; i < dadosLivros.length; i++) {
      if (dadosLivros[i][0] === idLivro) {
        linhaLivro = i + 1;
        tituloLivro = dadosLivros[i][1];
        if (dadosLivros[i][5] === 'Emprestado') {
          return { sucesso: false, mensagem: 'Não é possível dar baixa em livro emprestado.' };
        }
        if (dadosLivros[i][5] === 'Baixado') {
          return { sucesso: false, mensagem: 'Este livro já foi baixado.' };
        }
        break;
      }
    }
    
    if (linhaLivro === -1) {
      return { sucesso: false, mensagem: 'Livro não encontrado.' };
    }
    
    // Registrar baixa
    const sheetBaixas = getSheet_('Baixas');
    sheetBaixas.appendRow([
      gerarId_('BAI'),
      idLivro,
      tituloLivro,
      motivo,
      new Date().toLocaleDateString('pt-BR'),
      bibliotecario || 'Não informado'
    ]);
    
    // Atualizar status do livro
    sheetLivros.getRange(linhaLivro, 6).setValue('Baixado');
    
    return { sucesso: true, mensagem: 'Baixa registrada para: ' + tituloLivro };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// ---------- BIBLIOTECÁRIOS ----------

function cadastrarBibliotecario(dados) {
  try {
    const sheet = getSheet_('Bibliotecarios');
    const id = gerarId_('BIB');
    
    sheet.appendRow([
      id,
      dados.nome,
      dados.email || '',
      dados.telefone || '',
      new Date().toLocaleString('pt-BR')
    ]);
    
    return { sucesso: true, mensagem: 'Bibliotecário cadastrado! ID: ' + id };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

function listarBibliotecarios() {
  const sheet = getSheet_('Bibliotecarios');
  const dados = sheet.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    lista.push({
      id: dados[i][0],
      nome: dados[i][1],
      email: dados[i][2],
      telefone: dados[i][3]
    });
  }
  return lista;
}
