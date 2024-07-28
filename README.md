# mma_scorecards

Macros que sirven para rellenar e imprimir planillas de mma de la Federación Española de lucha en base a una lista de combates.

## Como usar

Descarga los ficheros Fightcard.xlsx y Scorecard_2024.xlsm en la **misma carpeta**. 

En el fichero Scorecard_2024.xlsm se encontrarán dos macros: 
"ProcessData" se encarga de leer la información de Fightcard y crear una carpeta con las planillas rellenas por cada combate en el listado.
"PrintAllFilesInFolder" se encarga de imprimir todas las planillas resultantes que han sido generadas por ProcessData.

Primero, edita Fightcard.xlsx con los combates apropiados para su evento.
Scorecard_2024.xlsm está preconfigurada para la impresión adecuada de las planillas para su uso en el evento. Se puede realizar ajustes en el caso de que sea necesario. 

Para lanzar las macros, abre Scorecard_2024.xlsm -> Desarrollador -> Macros
Escoge ProcessData y ejecuta la macro. Esto creará la carpeta Scorecard_Output y generará un fichero excel por cada combate listado en Fightcard. Puedes comprobar el resultado antes de ejecutar la siguiente macro.
A continuación, ejecuta PrintAllFilesInFolder para mandar a la cola de impresión todos los archivos generados que residen en la carpeta Scorecard_Output.

