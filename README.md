Macro VBA Excel PermutaciónSinRepetición

Descripción
Esta macro genera todas las permutaciones posibles sin repetición para un valor de tipo texto, numérico o alfanumérico. El resultado se muestra en una columna, comenzando desde una celda especificada por el usuario. La macro está limitada a un máximo de 9 caracteres, lo que genera hasta 362,880 permutaciones para evitar superar el límite de filas en Excel (1,048,576).

Características
Permite seleccionar una celda o ingresar manualmente el valor a permutar.
Genera permutaciones sin repetición, asegurando que cada combinación única se muestre una vez.
El mayor número de caracteres que se puede permutar sin exceder el límite de filas de Excel es 9, ya que 9! = 362,880 es el número máximo de permutaciones que pueden caber en una hoja de Excel. Las permutaciones de 10 caracteres generarían 3,628,800 combinaciones, lo que supera el límite de Excel.

Requisitos
Microsoft Excel 2007 o superior.
Habilitar macros en Excel.

Instrucciones de Uso
Abre el archivo de Excel y presiona ALT + F11 para abrir el editor de VBA.
Copia y pega el código de la macro en un nuevo módulo.
Cierra el editor de VBA y guarda el archivo como Libro habilitado para macros (*.xlsm).
Ejecuta la macro presionando ALT + F8, seleccionando la macro PermutacionSinRepeticion y presionando "Ejecutar".
La macro te preguntará si deseas seleccionar una celda que contenga el valor a permutar o ingresar manualmente.
Si seleccionas una celda, la macro leerá los primeros 9 caracteres de la celda seleccionada. Si ingresas las letras manualmente, asegúrate de no exceder los 9 caracteres.
A continuación, se te pedirá que selecciones la celda inicial donde se escribirán las permutaciones.
La macro generará todas las permutaciones y las insertará en la hoja de cálculo comenzando desde la celda especificada.

Limitaciones
Solo permite hasta 9 caracteres. Si se ingresan más de 9, solo se considerarán los primeros 9.
El tiempo de ejecución aumenta considerablemente con un mayor número de caracteres.
Para un conjunto de 9 caracteres, el tiempo de procesamiento puede ser largo debido a la gran cantidad de permutaciones.
