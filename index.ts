// versão base de typescript no Excel
function main(workbook: ExcelScript.Workbook) {
  // variaveis com listas, datas e contadores
  let total = 1
  let i = 1
  let lineValores: string[][] = []
  let lojasFaltantes: (string | number | boolean)[] = []
  let date = new Date().getDate()
  let dias: number
  // verifica quantas lojas existem na tabela
  let vendas = workbook.getActiveWorksheet()
  let lojas = vendas.getRange("C:C")
  for (let i = 1; lojas.getCell(i, 0).getText() != ""; i++) {
    total++
  }
  // tabelas e sheets
  let relatorio: ExcelScript.Worksheet
  let valores = vendas.getRange("C3:AI" + String(total + 1)).getTexts()
  // cria sheet Relatório
  if (workbook.getLastWorksheet().getName() != "Relatório") workbook.addWorksheet("Relatório")
  relatorio = workbook.getWorksheet("Relatório")

  // define quantos dias faltantes minimos para registrar a loja
  dias = 1
  valores.forEach((list) => {
    let i = 0
    for (const value of list) {
      if (i > 1) {
        if (value == "") {
          if (date - (i - 1) > dias) {
            lojasFaltantes.push(list[0])
            break
          }
        }
      }
      i++
    }
  })
  console.log(lojasFaltantes)
  // escreve no sheet Relatório as lojas que faltam informar
  relatorio.getRange("A:A").getCell(0, 0).setValue("Lojas Faltam Informar")
  for (const value of lojasFaltantes) {
    relatorio.getRange("A:A").getCell(i, 0).setValue(value)
    i++
  }
}
