Main script is ```Fill_Reichelt_Sheet.py```.

Call parameters:
* ```--service```: The uritools service URI
* ```--desktop_host```: LibreOffice desktop host, default=localhost
* ```--desktop_port```: LibreOffice desktop port, default=2002
* ```--sheet```: Path to LibreOffice sheet

Reichelt basket CSV is printed to stdout.

Note that LibreOffice must be running like

```
soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault
```

Do not use headless mode. The Sheet can be edited and must be saved manually.


https://pypi.org/project/pyoo/
