function fcIntegration() {
  const BANCO = ''; // Conector do banco
  const HOST = ''; // IP (0.0.0.0) ou HOST (seudominio.com.br)
  const PORTA = ''; // Porta para conexão
  const BANCODEDADOS = '' // Banco de dados desejado
  const USUARIO = ''; // Usuario
  const SENHA = ''; // Senha
  // const datahoje = new Date().getDay() + '/' + new Date().getMonth() + '/' + new Date().getFullYear();
  const ABA = 'FUNIL';

  const doc = SpreadsheetApp.getActiveSpreadsheet(); // Retorna a aba ativa
  const base = doc.getSheetByName(ABA); // Selecionamos a aba para limpar os dados
  const cell = base.getRange('A1'); // Vamos inserir os dados a partir da primeira célula

  // Criamos a conexão com o banco de dados
  const conn = Jdbc.getConnection('jdbc:' + BANCO + '://' + HOST + ':' + PORTA + '/' + BANCODEDADOS, USUARIO, SENHA);
  const stmt = conn.createStatement();
  const query = `SELECT content FROM kelibett_site.tpkb_formcraft_3_submissions ORDER BY created DESC LIMIT 10 OFFSET 20`;
  const rs = stmt.executeQuery(query); // Executamos a query para buscar em nosso banco de dados

  // Armazenamos os dados do banco em uma variável
  const registros = [];

  while (rs.next()) {
    for (let col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      const value = (rs.getString(col + 1).replace(/\\/g, ''));
      const entrie = JSON.parse(value);
      const fields = [];

      entrie.forEach((field) => {
        if (field.label == 'Text') return;
        fields.push({ label: field.label, value: field.value });
      });

      registros.push(fields);
    }
  }

  function getFirstEmptyRow() {
    let ct = 0;
    while (cell.offset(ct, 0).getValue() != '') {
      ct++;
    }
    return (ct);
  }

  // const rowValue = getFirstEmptyRow();
  // const s = registros.length - rowValue;

  // function addCell() {
  //   let range = base.getRange('Q2:Z2')
  //   for (i = 0; i <= s; i++) {
  //     range.insertCells(SpreadsheetApp.Dimension.ROWS);
  //   }

  //   console.log('addcell func')
  // }

  const usuariosAtuais = getFirstEmptyRow() - 1;
  const novosUsuarios = registros.length - usuariosAtuais;

  console.log(novosUsuarios)

  for (c = 0; c < novosUsuarios; c++) {
    base.insertRowAfter(0);
  }

  // Registramos os dados na planilha
  const labels = [];
  registros.forEach((registro, index) => {
    registro.splice(3, 1);
    registro.forEach((field, fieldIndex) => {
      // LABEL
      const label = field.label;

      if (!labels.find(name => name.labelName == label)) {
        cell.offset(0, fieldIndex).setValue(label.toUpperCase());
        labels.push({ labelName: label, labelIndex: fieldIndex });
      }

      // VALUE
      const valorFormated = formatValue(field.value);
      let labelIndexValue = 0;
      labels.find((name) => {
        if (name.labelName == label) {
          labelIndexValue = name.labelIndex;
        }
      });
      cell.offset(index + 1, labelIndexValue).setValue(valorFormated);
    });
  });

  rs.close();
  stmt.close();
  conn.close();
}


function formatValue(valor) {
  let str = valor.replace(/&atilde;/g, 'ã');
  str = str.replace(/&agrave;/g, 'à');
  str = str.replace(/&aacute;/g, 'á');
  str = str.replace(/&eacute;/g, 'é');
  str = str.replace(/&Eacute;/g, 'É');
  str = str.replace(/&uacute;/g, 'ú');
  str = str.replace(/&ocirc;/g, 'ô');
  str = str.replace(/&oacute;/g, 'ó');
  str = str.replace(/&ntilde;/g, 'ñ');
  str = str.replace(/&iacute;/g, 'í');
  str = str.replace(/&igrave;/g, 'ì');
  str = str.replace(/&ecirc;/g, 'ê');
  str = str.replace(/&amp;/g, '&');
  str = str.replace(/rn/g, '\n');
  str = str.replace(/&ccedil;/g, 'ç');

  return str
}
//Parte do código é da propria API do google, mas para termos um melhor aproveitamento da planilha temos mais linhas feitas a mão.
//Esse código também é uma parceria com "Eduardo Pinheiro(https://github.com/odraudep)"
