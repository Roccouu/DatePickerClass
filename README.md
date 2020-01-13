# DatePickerClass

> Un sencillo selector de fechas para Excel para ser usado con macros.

## Definición
DatePickerClass es un archivo de Clase VBA (.cls) con el que es fácil generar un sencillo Selector de Fechas mediante el uso de macros para hojas y formularios de Excel.

> Es necesario tener conocimientos básicos de programación de Clases VBA.

Trabaja bajo MS Excel versión 2007 o superior. No requiere instalación, sólo debe ser importardo a su Proyecto VBA.

Como DatePickerClass es una Clase VBA, debe instanciarse como un objeto y luego usar su interfaz de métodos y propiedades.

## Reconocimientos
Este proyecto fue inspirado en el trabajo desarrollado por [Andrés Rojas Moncada, Microsoft MVP](www.youtube.com/jarmoncada01) ([Sitio web](https://www.excelhechofacil.com/)): Calendario con UserForm para su Versión 1.0 de 20 de julio de 2015 y usa el núcleo de su algoritmo. Muchas gracias Andrés por todo tu talento.

## Modo de uso
  1.  [Descargue el archivo DatePickerClass.cls](https://github.com/Roccouu/DatePickerClass/tree/master/project-dev/DatePickerClass.cls).
  2.  Cree un nuevo libro habilitado para macros de Excel.
  3.  Abra el Editor de Proyectos VBA con Ctrl+F11
  4.  Vaya al menú ***Archivo > Importar Archivo*** o presione ***Ctrl+M*** y luego busque su archivo DatePickerClass.cls descargado e impórtelo al Proyecto VBA.
  5.  Puede crear un nuevo módulo y un nuevo formulario donde se utilizará ***DatePickerClass*** (también puede descargar [datepickerexampleuse.xlsm](https://github.com/Roccouu/DatePickerClass/tree/master/project-dist/datepickerexampleuse.xlsm) para ver algunos ejemplos prácticos y fáciles de usar ***DatePickerClass***)
  6.  La Clase tiene una Propiedad y dos Métodos:

      1. La propiedad ```GetDATE``` devuelve la fecha seleccionada con el objeto ***DatePickerClas***. Su valor predeterminado es la fecha actual del sistema.
      2. Con el Método ```DatePickerAdd(clFrm, [clFirstDay])``` puede agregar el Control de selector de fecha creado internamente en la Clase en el UserForm creado previamente en su Proyecto VBA. Este método requiere dos parámetros, el primero es el UserForm y el segundo es una constante de VBA tipo día, opcional (VBA.vbMonday o VBA.vbSunday o el primer día de la semana en su calendario), su valor predeterminado es ```VBA.vbMonday```.
      3. Con el Método ```DatePickerUse([clJustForm], [clControl], [clAlign], [clBaseColor], [clMsgBox])``` puede usar el objeto de dos formas posibles: primero en un UserForm, para esta tarea, el parámetro ```clControl``` requiere un objeto de control como **TextBox, Label, Button o ComboBox**, este control recibirá la fecha seleccionada por el usuario. La segunda forma permitirá el uso de UserForm como un simple Selector de Fechas en la hoja de Excel que usted elija, esta parte es muy buena; para esta tarea, ```clJustForm``` debe configurarse en ```True``` y nada más. El parámetro ```clAlign``` requiere un tipo de dato ```String```: "R" (de "Derecha"), "L" (de "Izquierda") o "C" (de "Centro") para decir al objeto DatePicker se alínee respecto ```clControl```, el valor predeterminado es "L". ```clBaseColor``` es un Long opcional para establecer el color de estilo de **DatePicker**, puede ser el resultado de la función nativa ```VBA.RGB(RR, GG, BB)``` de VBA. Finalmente, ```clMsgBox``` es un valor booleano opcional, predeterminado: ```False```, este parámetro le dice a **DatePickerClass**, muestra un ```VBA.msgbox``` con la fecha seleccionada por el usuario.
      7. ¡Disfruta de **DatePickerClass**!

## Colaborar en GitHub:
El código fuente de **DatePickerClass** está en: [el directorio project-dev](https://github.com/Roccouu/DatePickerClass/tree/master/project-dev/DatePickerClass.cls) del repositorio oficial.

Tan pronto como se descargue, puede colaborar con mejoras en el Sistema siempre bajo el respeto de [Términos de licencia](https://github.com/Roccouu/DatePickerClass/blob/master/LICENSE), [El Código de Conducta](https://github.com/Roccouu/DatePickerClass/blob/master/CODE_OF_CONDUCT.md) y los [Términos de Contribución](https://github.com/Roccouu/DatePickerClass/blob/master/CONTRIBUTING.md).

## Sitio Web

[DatePickerClass](https://roccouu.github.io/DatePickerClass/docs/index.html)

## Tutorial

[Tutorial DatePickerClass](https://roccouu.github.io/DatePickerClass/docs/index.html#/tutorial)

## Documentación

[Documentación DatePickerClass](https://roccouu.github.io/DatePickerClass/index.html#/docs/index.html#/documentation)

## Contribución

Vea las [Guías de CONTRIBUCIÓN](https://github.com/roccouu/DatePickerClass/CONTRIBUTING.md)

## English Readme

[README-EN.md](https://github.com/roccouu/DatePickerClass/blob/master/README-EN.md)

## Licencia

[MIT](https://github.com/roccouu/DatePickerClass/blob/master/LICENSE) © | [Roccou](https://twitter.com/_roccou) | 2020