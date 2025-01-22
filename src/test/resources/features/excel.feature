Feature: Manipulacion de archivos excel

  Scenario: Convertir de excel a json
  Given def excel = Java.type("util.ExcelUtil")
    * def jsonVentas = excel.toJson("src/test/resources/features/ventas.xlsx")
    * print jsonVentas
    * print jsonVentas.data[0].Cliente

  Scenario: Convertir de json a excel
  Given def logisticaData = read('envios.json')
    * print logisticaData.headers
