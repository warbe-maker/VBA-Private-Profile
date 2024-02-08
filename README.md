## VBA Private Profile File Services
Services provided in a VBA manner which simplifies and unifies read/write of _Private Profile_ files, for example:  
- Read Value: `<value> = Value(<value_name>[, <section_name>][, <file_name>])`
- Write Value: `Value(<value_name>[, <section_name>][, <file_name>]) = <value>`. 
A unique experience thereby: All sections and values are written in ascending order which eases reading of large files.

### The _Value_ service
The service has the following named arguments:

| Argument        | Explanation |
|-----------------|-------------|
| _file\_name_    | String expression, specifies the _PrivateProfile_ file by its full name, defaults to `ThisWorkbook.FullName` with the file extension replaced by `.dat`. The file name may alternatively be explicitly specified once by the  _FileName_ property. Once either way specified, it becomes the default for all subsequent methods and properties when omitted.|
| _section\_name_ | String expression, may be specified once with the _SectionName_. Once either way specified, it becomes the default for all subsequent methods and properties when omitted.|
| _value\_name_   | |

### Other services

| Type     | Service    | Description |
|----------|---------------|-------------|
| Property |_FileName_&nbsp;r/w | Specifies a _Private Profile_ file full name. |
|          |                    | Returns the current valid _Private Profile_ file's full name |
| Property |_Section_&nbsp:r/w  | Returns the current specified Section name |
|          |                    | Specifies the Section name valid for all subsequent methods and properties until another section name is specified.|
| Property |_Value_&nbsp;r/w    | Reads from a _Private Profile_ file a value with a provided value-name from a provided section.|
|          |                    | Writes to a Private Profile File a value with a provided value-name into a provided section.|
| Method   |_NamesRemove_       | Removes provided value names, in a given _Private Profile file, optionally only in provided specific sections, when none are provided, in all sections.|
| Method   |_SectionExists_     | Returns TRUE when a given section exists in the current valid _Private Profile_ file.|
| Method   |_SectionNames_      | Returns a Dictionary of all section names.|
| Method   |_SectionsCopy_      | Copies sections, provided as comma delimited string of section names from a source _Private Profile_ file into a target _Private Profile_ file, optionally merged.|
| Method   |_SectionsRemove_    | Removes the sections provided as a comma delimited string, whereby sections not existing are ignored.|
| Method   |_ValueNameExists_   | Returns TRUE when a given value-name exists in a provided _Private Profile_ file's section.|
| Method   |_ValueNameRename_   | Replaces an old value name with a new one either in sections provided as a comma delimited string or in all sections when none are provided.|
| Method   |_ValueNames_        | Returns a Dictionary of all value-names within sections provided as a comma delimited string or of all sections when no sections are provided, with the value name as the key and the value as item. Because any duplicate names are ignored, the value will be the value of the first found.|

## Installation
1. Download and import [clsPrivProf.cls][1] to your VB project.
2. In the VBE add a Reference to _Microsoft Scripting Runtime_ and _Microsoft VBScript Regular Expression 5.5_

### Usage
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
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated. The module is available in a dedicated [Workbook][3] (public GitHub repository) which includes a complete regression test of all services.

[1]:https://github.com/warbe-maker/VBA-Private-Profile/blob/main/source/clsPrivProf.cls
[2]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[3]:https://github.com/warbe-maker/VBA-Private-Profile