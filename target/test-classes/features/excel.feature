Feature: Manipulacion de archivos excel

  Background:
    Given def excelUtil = Java.type("util.ExcelUtil")

#  Scenario: Convertir de excel a json
#  Given def jsonVentas = excelUtil.toJson("src/test/resources/features/ventas.xlsx")
#    * print jsonVentas
#    * print jsonVentas.data[0].Cliente

  Scenario: Convertir de json a excel
  Given def logisticaData = read('envios.json')
    * excelUtil.toExcel(karate.toString(logisticaData))
#    * excelUtil.toExcel(logisticaData)
    * print logisticaData.headers
