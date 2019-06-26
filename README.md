XLSX
[![Build Status](https://travis-ci.org/thrust-bitcodes/xlsx.svg?branch=master)](https://travis-ci.org/thrust-bitcodes/xlsx) [![GitHub release](https://img.shields.io/github/release/thrust-bitcodes/xlsx.svg)](https://github.com/thrust-bitcodes/xlsx/releases)
===============

XLSX é um *bitcode* de modelagem de xlsx para [thrust](https://github.com/thrustjs/thrust) que utiliza POI-XSSF como mecanismo de manipulação.

# Instalação

Posicionado em um app [thrust](https://github.com/thrustjs/thrust), no seu terminal:

```bash
thrust install xlsx
```

## Tutorial de escrita

```javascript
//Realizamos o require do bitcode
var xlsx = require('xlsx')

var rows = [{
    id: 1,
    nome: 'Zé',
    dt_nasc: new Date(),
    salario: 1122.5
}, {
    id: 2,
    nome: 'Jão',
    dt_nasc: new Date(),
    salario: 965.7
}, {
    id: 3,
    nome: 'Maria',
    dt_nasc: new Date(),
    salario: 1100.0
}];

/**
 Todo o objeto de metadados é opcional
 Os possiveis estilos de célula são:
  - horizontalAlignment: <String> - left, center, right, fill, justify
  - fontName: <String> Nome da fonte
  - fontSize: <Number> tamanho da fonte
  - fontColor: <String> Cor da Fonte.
  Acesse https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html para lista de opções
  - size: <Number> Tamanho da coluna. É utilizado quando a opção autoSize possui o valor `false`,
  - backgroundColor: <String> Cor de fundo da célula. Acesse https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html para lista de opções
  - unlocked: <Boolean> Seta a célula como desbloqueada. Só funciona quando uma senha é atribuida para a planilha.
  - verticalAlignment: <String> top, center, bottom
  - formula: <Boolean> Quando ativo pega o valor da célula e aplica como uma fórmula
  - bold: <Boolean> se será negrito ou não.
  - italic: <Boolean> se será itálico ou não
  - striked: <Boolean> se terá strikethrough
  - underline: <Boolean> se terá underline
  - doubleUnderline: <Boolean> se terá double underline.

  Ambas as propriedades headerStyle, style, e objetos de coluna, podem ter estes estilos, sendo que cada um tem precedência sobre o outro, de forma que seja possível realizar uma configuração geral para todos e algumas específicas.

  Precedência de estilos:
   - HEADER: headerStyle, column, style
   - DEMAIS: column, style

 Os objetos de coluna como dito, pode ter as mesma propriedades de estilo citadas acima e além deles:
  - description: <String> - Se presente, irá alterar o header desta coluna
  - format: <String> - Determina a formação da célula
  - type: <String> - Determina os tipos de formatação de uma célula, caso format não tenha sido informado, podendo ser:
    - currency = 'R$ #,##0.00'
    - time = 'HH:MM'
    - date = 'DD/MM/yyyy'
    - datetime = date + time
*/
var metadata = { //Opcional
    hasHeader: true, //Se deverá ser criada uma linha de header
    asByteArray: true, //Se o resultado é um byte[] ou o workbook do POI
    autoSize: true, //Se deve fazer o ajuste automatico de largura das celulas
    orderByColumnConfig: true, //se deve utilizar ordenação de acordo com o parametro columns 
    headerStyle: { //Estilo que será aplicado ao header, caso exista
        bold: true,
        horizontalAlignment: 'center',
        fontName: 'Colibri'
    },
    rowStyle: { // Estilos específica da linha
        height: 1000 // Altura das linhas da planilha
    }
    style: { //Estilo padrão que será aplicado a todas as celulas
        fontName: 'Courrier New',
    },
    columns: { //Configurações específicas de uma coluna
        id: {
            description: 'Código', //Se presente irá usar esta string como header
            italic: true, //estilo da coluna
        },
        nome: {
            doubleUnderline: true
        },
        dt_nasc: {
            description: 'Dt Nasc.',
            striked: true,
        },
        salario: {
            type: 'currency',
        },
        somatoria: {
            formula: true
        }
    },
    sheetProperties : { // Configurações específicas da sheet
      password: '123', // Senha, que ao ser setada bloqueia toda a planilha. Utilize a propriedade `unlocked` para desbloquear as células de uma coluna
      name: 'Planilha 1' // nome da aba
    }
};

//Para gerar uma planilha
var bytes = xlsx.create(rows, metadata);

//Para salvar os bytes da planilha em um arquivo no disco
xlsx.writeBytesToFile(bytes, './planilha.xlsx');
```

## Tutorial de leitura

```javascript
//Realizamos o require do bitcode
var xlsx = require('xlsx');

var metadata = { //Opcional
    hasHeader: true //Determina se a sua planilha possui header
}

/*
 O primeiro argumento pode ser:
  -  <String> Path do arquivo
  - <java.io.InputStream> stream da planilha
  - <byte[]> bytes da planilha
*/
var jsonPlanilha = xlsx.read('./planilha.xlsx', metadata);
```
