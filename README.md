# SimpleExcelFile
Este repositorio genera un reporte basado en el formato que
pide la empresa Internacional de Sistemas, este reporte
sirve para ver la horas de trabajo realizadas. el siguiente es un ejemplo del archivo enviado:

![alt-image](https://i.imgur.com/ggr6Rbl.png "Reporte de Horas")

## Requisitos
- Python 3.x
- Paquetes Python:
    - Datetime
    - openpyxl = 2.5.4
    - email
    - calendar

## Estructura
El projecto se estructura de la siguiente manera:

- SimpleExcelFile
    - README.md
    - SimpleExcelFile
        - __init__py
        - DateData.py
        - EmailDetails.py
        - ExcelFile.py
        - main.py
    - Test
        - test_file.py
        - test_mail.py
        
## Ejecución
- main.py
    - dentro del archiv main hay que modificar ciertos parametros para que el archivo
    correo sean enviados de forma correcta, los parametros a modificar son:
        - __passwd__, Contraseña de la cuenta de correo por donde seran enviados los archivos.
        - __fromaddr__, Cuenta de correo electronico desde la cual seran enviados los archivos.
        - __toaddr__, Cuentas de correo a las cuales les seran enviados los archivos, delimitadas por comas.
        - __message__, Default
            - __nombre__, nombre de la persona.
        - __message__, Mensaje personalizado
    - Ejecutar el archivo con el siguiente codigo(cmd, bash):
    ```Bash
    python3 main.py
    # Se generara el archivo y se eviara el correo segun se especifique.
    ```
## Test Practicos
- __test_file.py__, este archivo genera el archivo segun la fecha especificada.
    - Modificar la fecha del archivo, segun se requiera.
    - Se ejecuta de la siguiente manera(cmd, bash).
    ```Bash
    python3 test_file.py
    ```
- __test_mail.py__, este test envia un correo electronico
    -Se trata de un archioo que ba gtra a 100 peroms
    - Se ejecuta de la siguiente manera(cmd, bash):
    ```Bash
    python3 test_mail.py
    ```        
