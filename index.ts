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
  // cria sheet Relatório e verifica se já existe
  let relSheet = workbook.getWorksheet("Relatório")
  if (relSheet == undefined) { workbook.addWorksheet("Relatório") } 
  else {relSheet.delete(); workbook.addWorksheet("Relatório")}
  relatorio = workbook.getWorksheet("Relatório")

  // define quantos dias faltantes minimos para registrar a loja
  dias = 10
  valores.forEach((list) => {
    let valido = true
    let validoComeco = true
    let i = 0
    let diasInicio = 0
    // verifica quais lojas podem ser cobrados mesmo que tenham dias iniciais faltantes
    // Ex: A loja começou no dia 15/12/2025 e começou a informar as vendas apenas agora, nisso os 15 dias anteriores não devem ser considerados
    for (const value of list) {
      if (i > 1) {
        if(value == "" && validoComeco) {
          diasInicio++
        } else if (value == "" && !validoComeco) {
          if (date - (i - 1) - diasInicio > dias && list[i + 1] == "") {
            lojasFaltantes.push(list[0])
            break
          }
        } else {
          validoComeco = false
        }
      }
      i++
    }
  })
  // escreve no sheet Relatório as lojas que faltam informar
  relatorio.getRange("A:A").getCell(0, 0).setValue("Lojas Faltam Informar")
  for (const value of lojasFaltantes) {
    relatorio.getRange("A:A").getCell(i, 0).setValue(value)
    i++
  }
}
