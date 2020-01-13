# DatePickerClass

> Excel simple date picker for use with macros.

## Definition
DatePickerClass is a VBA Class file (.cls) with wich is easy to generate a simple Date Picker for use with macros on Excel Sheets and VBA Forms.

> It is necessary to have basic knowdlements about VBA Class programation.

It works under MS Excel version 2007+. No installation required, just import to your VBA Project.

As DatePickerClass is a VBA Class, it needs to be instanciated as an Object and then use its methods and properties interface.

## Acknowledgments
This project was inspired by the work developed by [Microsoft MVP, Andrés Rojas Moncada](www.youtube.com/jarmoncada01) ([web page](https://www.excelhechofacil.com/)) Calendar in with UserForm for Version 1.0 of July 20, 2015 and uses the core of its algorithm. Thank you very much Andrés for all your talent.

## Use Mode
  1.  [Download DatePickerClass.cls](https://github.com/Roccouu/DatePickerClass/tree/master/project-dev/DatePickerClass.cls) file.
  2.  Create a new Excel enabled for macros workbook.
  3.  Open the VBA Editor with Ctrl+F11
  4.  Go to ***File > Import File*** menu or press ***Ctrl+M*** and then search your DatePickerClass.cls downloaded file and import it to the VBA Project.
  5.  You can create a new Module and a new Form where ***DatePickerClass*** will be used (also you can download the [datepickerexampleuse.xlsm](https://github.com/Roccouu/DatePickerClass/tree/master/project-dist/datepickerexampleuse.xlsm) file to see a practical and easy ways of use ***DatePickerClass***)
  6.  The Class has one Property and two Methods:

      1.  Property ```GetDATE``` returns the date selected with the DatePickerClas object. Its default value is the current system date.
      2.  With the Method ```DatePickerAdd(clFrm, [clFirstDay])``` you can add the Date Picker Control created internally in the Class at the UserForm previously created in your VBA Project. This method requires two parameters, first is the UserForm and the second is and optional VBA day constant (VBA.vbMonday or VBA.vbSunday or the first day of week on your calendar), its default value is ```VBA.vbMonday```.
      3.  With the Method ```DatePickerUse([clJustForm], [clControl], [clAlign], [clBaseColor], [clMsgBox])``` you can use the object on two possible ways: first into a UserForm, to this task, the parameter ```clControl``` requires an Control Object as a TextBox, Label, Button or a ComboBox, this control will receive the date selected by the user. The Second way will allow the use of the UserForm as a simple DatePicker on Excel WorkSheets, it is very cool, to this task, ```clJustForm``` must to be setted at ```True``` and nothing else.
      Parameter ```clAlign``` require a data type String: "R" (of "Right"), "L" (of "Left") or "C" (of "Center") to say DatePicker Object align respect ```clControl```, default value is "L". ```clBaseColor``` is a optional Long to set the style color of DatePicker, it can be the result of VBA.RGB(RR,GG,BB) native function of VBA. Finally, ```clMsgBox``` is a optional Boolean, default: False, this parameter says to DatePickerClass, show a VBA.msgbox with the date selected by the user.
  7.  Enjoy DatePickerClass!


## Collaboration on GitHub:
**DatePickerClass** source code is in: [project-dev folder](https://github.com/Roccouu/DatePickerClass/tree/master/project-dev/DatePickerClass.cls) into this Official repository.
As soon it is downloaded, you can collaborate with improvements to the System always under respect of [License terms](https://github.com/Roccouu/DatePickerClass/blob/master/LICENSE), [Code of conduct](https://github.com/Roccouu/DatePickerClass/blob/master/CODE_OF_CONDUCT.md) and the [Contribution terms](https://github.com/Roccouu/DatePickerClass/blob/master/CONTRIBUTING.md).

## Website

[DatePickerClass](https://roccouu.github.io/DatePickerClass/docs/index.html)

## Tutorial

[DatePickerClass tutorial](https://roccouu.github.io/DatePickerClass/docs/index.html#/tutorial)

## Documentation

[DatePickerClass Docs](https://roccouu.github.io/DatePickerClass/index.html#/docs/index.html#/documentation)

## Contributing

See the [CONTRIBUTING Guidelines](https://github.com/roccouu/DatePickerClass/CONTRIBUTING.md)

## License

[MIT](https://github.com/roccouu/DatePickerClass/blob/master/LICENSE) © | [Roccou](https://twitter.com/_roccou) | 2020