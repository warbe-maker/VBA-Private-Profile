## VBA _Private Profile_ file services
Simplify and unify _Private Profile_ file services by supporting:
- section separation (default)
- file header and footer
- section and value comments
- content in ascending order (on the fly)
- any number of _Private Profile_ files in one class instance [^1] 

The following is an example of the simplification:  
- Value read: `<value> = Value(<value_name>, <section_name>[, <file_name>])`
- Value write: `Value(<value_name>, <section_name>[, <file_name>]) = <value>`. 
A unique experience thereby: All sections and values are written in ascending order which eases reading of large cfg-, ini-, and other kinds of _Private Profile_ files.

### The _Value_ service
The service has the following named arguments:

| Argument        | Explanation |
|-----------------|-------------|
| _name\_value_   | String expression, name of the value in the _Private Profile_ file.|
| _name\_section_ | String expression, optional, specifies the _Section_ for the value, defaults to the section specified through the property _Section_.|
| _name\_file_    | String expression, optional, specifies the full name of a _Private Profile_ file, defaults to file specified through the property _FileName_ when omitted.|

### Other services

| Service            | Type     | Description |
|--------------------|---------------|-------------|
|_FileName_&nbsp;r/w | Property | Returns/specifies a _Private Profile_ file's full name. When none has been specified the file name defaults to _ThisWorkBook.Path & "\PrivateProfile.dat". |
|_Section_&nbsp;r/w  | Property | Returns/specifies the Section name used throughout subsequent services when the section argument is omitted. |
|_Value_&nbsp;r/w    | Property | Reads from a _Private Profile_ file a value by a provided value-name in a provided section (defaults to the _Section_ property when omitted). When the file name is omitted it defaults to the name specified by the _FileName_ property.|
|          |                    | Writes to a _Private Profile_ file a value by a provided value-name in a provided section (defaults to the _Section_ property when omitted). When the file name is omitted it defaults to the name specified by the _FileName_ property.|
|_NamesRemove_       | Method   | Removes provided value names, in a given _Private Profile file, optionally only in provided specific sections, when none are provided, in all sections.|
|_SectionExists_     | Method   | Returns TRUE when a given section exists in the current valid _Private Profile_ file.|
|_SectionNames_      | Method   | Returns a Dictionary of all section names.|
|_SectionsRemove_    | Method   | Removes the sections provided as a comma delimited string, whereby sections not existing are ignored.|
|_ValueNameExists_   | Method   | Returns TRUE when a given value-name exists in a provided _Private Profile_ file's section.|
|_ValueNameRename_   | Method   | Replaces an old value name with a new one either in sections provided as a comma delimited string or in all sections when none are provided.|
|_ValueNames_        | Method   | Returns a Dictionary with all value names a _Private Profile_ file with the value name as the key and the value as the item, of all sections if none is provided or those of a provided section's name. When the file name is omitted it defaults to the name specified by the _FileName_ property.<br>***Note:*** The returned value-names are distinct names! I.e. when a value exists in more than one section it is still one distinct value-name.|
| _SectionSeparation_ | Property | Boolean expression, default to True, separates sections by an empty line to improve readability.|

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
    
    '~~ Write a value
    Cfg.Value("AnyValueName", "AnySection") = "Any value"
    
    '~~ Read a value
    myvalue = Cfg.Value("AnyValueName", "AnySection") 
```

> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time is prepared to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][2] for more details.

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding rules and guidelines will be appreciated. The module is available in a dedicated [Workbook][3] (public GitHub repository). This Workbook also provides a complete regression test covering all public services (methods and properties).

[^1]: Though possible it will be more elegant to use a class instance for each individual _Private Profile_ file, provide the file by the _FileName_ property once and omit it in all other properties and methods.

[1]:https://github.com/warbe-maker/VBA-Private-Profile/blob/main/source/clsPrivProf.cls
[2]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[3]:https://github.com/warbe-maker/VBA-Private-Profile