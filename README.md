## VBA _Private Profile_ file services
Simplifying and unifying _Private Profile_ file services by providing all of them in a VBA manner, thereby allowing and considering file, section, and value headers and also a file footer. The following is an example of the simplification:  
- Value read: `<value> = Value(<value_name>[, <section_name>][, <file_name>])`
- Value write: `Value(<value_name>[, <section_name>][, <file_name>]) = <value>`. 
A unique experience thereby: All sections and values are written in ascending order which eases reading of large cfg-, ini-, and other kinds of _Private Profile_ files.

### The _Value_ service
The service has the following named arguments:

| Argument        | Explanation |
|-----------------|-------------|
| _name\_value_   | String expression, name of the value in the _Private Profile_ file.|
| _name\_section_ | String expression, optional, specifies the _Section_ for the value, defaults to the section specified through the property _Section_.|
| _name\_file_    | String expression, optional, specifies the full name of a _Private Profile_ file, defaults to file specified through the property _FileName_ when omitted.|

### Other services

| Type     | Service    | Description |
|----------|---------------|-------------|
| Property |_FileName_&nbsp;r/w | Returns/specifies a _Private Profile_ file's full name. When none has been specified the file name defaults to _ThisWorkBook.Path & "\PrivateProfile.dat". |
| Property |_Section_&nbsp;r/w  | Returns/specifies the Section name used throughout subsequent services where the section argument is omitted. |
| Property |_Value_&nbsp;r/w    | Reads from a _Private Profile_ file a value with a provided value-name from a provided section.|
|          |                    | Writes to a _Private Profile_ file a value with a provided value-name into a provided section.|
| Method   |_NamesRemove_       | Removes provided value names, in a given _Private Profile file, optionally only in provided specific sections, when none are provided, in all sections.|
| Method   |_SectionExists_     | Returns TRUE when a given section exists in the current valid _Private Profile_ file.|
| Method   |_SectionNames_      | Returns a Dictionary of all section names.|
| Method   |_SectionsCopy_      | Copies sections, provided as comma delimited string of section names from a source _Private Profile_ file into a target _Private Profile_ file, optionally merged.|
| Method   |_SectionsRemove_    | Removes the sections provided as a comma delimited string, whereby sections not existing are ignored.|
| Method   |_ValueNameExists_   | Returns TRUE when a given value-name exists in a provided _Private Profile_ file's section.|
| Method   |_ValueNameRename_   | Replaces an old value name with a new one either in sections provided as a comma delimited string or in all sections when none are provided.|
| Method   |_ValueNames_        | Returns a Dictionary of all value-names within sections provided as a comma delimited string or of all sections when no sections are provided, with the value name as the key and the value as item. Because any duplicate names are ignored, the value will be the value of the first found.|
| Method  | _Reorg_             | Explicit reorganization of a _Private Profile_ file which means that all sections and value names are ordered in ascending sequence thereby considering file, section, and value headers, and a possible file footer.|      

## Installation
1. Download and import [clsPrivProf.cls][1] to your VB project.
2. In the VBE add a Reference to _Microsoft Scripting Runtime_ and _Microsoft VBScript Regular Expression 5.5_

## Usage (example)
In a Standard Module:  
`Public Dim Cfg As New clsPrivProf`

In any other module (for example):  
```vb  
    '~~ Specifiy the Private Profile's full file name and a section name which then can be
    '~~ omitted with any subsequent service call
    Cfg.FileName = ThisWorkbook.Path & "\App.cfg"
    Cfg.Section = "AnySectionName"
    
    '~~ Write a value
    Cfg.Value("AnyValueName") = "Any value" file
    
    '~~ Read a value
    myvalue = Cfg.Value("AnyValueName") 
```

> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][2] for more details.

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding rules and guidelines will be appreciated. The module is available in a dedicated [Workbook][3] (public GitHub repository). This Workbook also provides a complete regression test including all public services.

[1]:https://github.com/warbe-maker/VBA-Private-Profile/blob/main/source/clsPrivProf.cls
[2]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[3]:https://github.com/warbe-maker/VBA-Private-Profile