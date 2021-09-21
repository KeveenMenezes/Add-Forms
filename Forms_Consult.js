function onOpen()
{
    let ui=SpreadsheetApp.getUi();
    ui.createMenu("Criações de formulário").addItem("Executar", "formulario").addToUi();
}

function formulario() {
  let c = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getMaxRows();
  for(let a = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange("AJ1").getValue();a<=c;a++){
    let numb = a.toString();
    let celM = ("M" + numb);
    let celAJ = ("AJ" + numb);
    let celBQ = ("BQ" + numb);
    let celBR = ("BR" + numb);
    let sheetM = null; // id
    sheetM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange(celM).getValue();
    let sheetAJ = null; // link do Formulario de inscrição
    sheetAJ = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange(celAJ).getValue();

      
      if( sheetAJ == "" &&  sheetM != "" ){

        // Estrutura do Forms
        let form = FormApp.create(sheetM);

        form.setRequireLogin(false);
        form.setProgressBar(true);
        form.setConfirmationMessage("Muito Obrigado pela sua colaboração. Ate breve !!!");

        form.setTitle('Formulário de Reação').setDescription("Bem-vindxs ao questionário virtual de avaliação de Reação. Por gentileza leia atentamente cada pergunta antes de responder. Preparamos algumas orientações para facilitar o processo de realização do questionário:\n\nSua participação é muito importante para nós. Respondendo esse breve questionário, você nos ajuda a entender o quanto essas experiências contribuíram para seu desenvolvimento profissional, para suas inspirações e a melhorar os próximos treinamentos.\n\nFique tranquilo suas respostas são confidenciais e serão tratadas junto a dos demais participantes.Em caso de quaisquer dúvidas, gentileza entrar em contato com o analista de qualidade do seu regional.");
        form.addMultipleChoiceItem()
        .setTitle('Selecione o ID do formulário:')
        .setChoiceValues([sheetM]).setRequired(true);
        form.addPageBreakItem()
        .setTitle('Avaliação de Reação').setHelpText("Essa seção seguirá com perguntas para verificar se suas expectativas foram atendidas após a realização do treinamento. Seu feedback é muito importante, pois contribui para a melhoria contínua dos nossos treinamentos. Por gentileza atribua uma nota de 1 à 10 para cada quesito.\n\n\nConteúdos, de 1 a 10 classifique. ");

        form.addScaleItem()
        .setTitle('1) Relação do conteúdo com o dia-a-dia de trabalho:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('2) Nível de profundidade do conteúdo:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('3) Qualidade das técnicas e exercícios aplicados:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('4) Carga horária e dinamismo do curso:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('5) Cumprimento dos objetivos:')
        .setBounds(0, 10).setRequired(true);

        form.addPageBreakItem()
        .setTitle('Facilitador, de 1 a 10 classifique.')
        form.addScaleItem()
        .setTitle('6) Conhecimento do tema:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('7) Clareza e objetividade ao explicar:  ')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('8) Estímulo à participação do grupo:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('9) Cordialidade e atenção com os participante:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('10) Pontualidade e aproveitamento do tempo')
        .setBounds(0, 10).setRequired(true);

        form.addPageBreakItem()
        .setTitle('Recursos e Conexão, de 1 a 10 classifique. ')
        form.addScaleItem()
        .setTitle('11) Localização / Acesso :')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('12) Infraestrutura físico e/ou virtual:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('13) Recursos técnicos utilizados, equipamentos e conexão:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('14) Conteúdo do Material didático compatível com os objetivos:')
        .setBounds(0, 10).setRequired(true); form.addScaleItem()
        .setTitle('15) Layout do Material didático e ambiente de aprendizagem:')
        .setBounds(0, 10).setRequired(true); 

        form.addPageBreakItem()
        .setTitle('Recomendações, Expectativa e Satisfação Geral do Curso')
        form.addGridItem().setTitle("16) Você recomendaria este curso a um colega?")
        .setRows([''])
        .setColumns(["Sim","Não"]).setRequired(true);
        form.addGridItem().setTitle("17)Você recomendaria este consultor / instrutor a um colega?")
        .setRows([''])
        .setColumns(["Sim","Não"]).setRequired(true);
        form.addGridItem().setTitle("18)Este curso e as atividades realizadas, atenderam as suas expectativas?")
        .setRows([''])
        .setColumns(["Sim","Não"]).setRequired(true);
        form.addScaleItem()
        .setTitle('19) De 0 a 10, qual nota você daria para este curso?')
        .setBounds(0, 10).setRequired(true); 
        form.addTextItem().setTitle("20) Por favor, deixe a sua mensagem. Gostaríamos de saber no que podemos melhorar para colaborar com o seu desenvolvimento.");
        
        let sheetRespostas = SpreadsheetApp.create(sheetM);
        form.setDestination(FormApp.DestinationType.SPREADSHEET, sheetRespostas.getId());
        let ultcell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange("AJ1").setValue(numb);

        

// Obter links
        let public = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange(celAJ).setValue(form.shortenFormUrl(form.getPublishedUrl().toString()));
        let edit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange(celBQ).setValue(form.getEditUrl());
        let data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cronograma Geral').getRange(celBR).setValue("https://docs.google.com/spreadsheets/d/"+form.getDestinationId());
      }  
  }
}
