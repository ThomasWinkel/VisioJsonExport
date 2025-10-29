# VisioJsonExport
JSON export that follows the Visio object model.
Allows you to process a document without a Visio application.

The export does only contain specific data like:
- User defined cells
- Properties
- Connections

The export does not contain common data like:
- Regular ShapeSheet cells (positions, colors, etc.)

In future we can extend data as reuired.

## Usage
### Export in .NET
For export we provide a NuGet package:
https://www.nuget.org/packages/Geradeaus.VisioJsonExport
```C#
Geradeaus.VisioJsonExport.ExportHandler exportHandler = new Geradeaus.VisioJsonExport.ExportHandler(Globals.ThisAddIn.Application.ActiveDocument);
exportHandler.Parse();
exportHandler.Export(@"C:\Temp\VisioExport.json");
```

### Process in Python
For processing in Python we provide a package on PyPi:
https://pypi.org/project/visio-json-export
```python
import visio_json_export

visio = visio_json_export.load_file(r'C:\Temp\VisioExport.json')

for page in visio.document.pages.values():
    for shape in page.shapes.values():
        for row_name, user in shape.user_rows.items():
            print(page.name + ' -> ' + shape.name + ' -> ' + row_name + ' = ' + user.value)
        for prop in shape.prop_rows.values():
            print(page.name + ' -> ' + shape.name + ' -> ' + prop.label + ' = ' + prop.value)
```