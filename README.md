# SICOFIN

El Sistema Internet de Contabilidad Financiera (SICOFIN) fue construido a la medida de las necesidades 
del Banco Nacional de Obras y Servicios Públicos, S.N.C. (Banobras), en México entre los años 2000 y 2002.

Este repositorio contiene el código de dicho sistema.

El proyecto fue desarrollado en Visual Basic 6.0, bajo páginas HTML basadas en Active Server Pages (ASP 1) y JavaScript, y utiliza Oracle como servidor de base de datos.

# Contenido

1. **Binaries**  
   Los archivos DLL del código VB en su última versión. Se incluyen en el repositorio en virtud de que puede ser complicado 
   regenerar las DLLs por la antigüedad de las herramientas con las que fue desarrollado.

2. **Configuration Data**  
   Archivos de configuración del sistema. Hay archivos XML y un par de archivos .reg cuyo contenido se guarda en el registro de Windows Server.

3. **Database Scripts**  
   Scripts con la definición de tablas, stored procedures, y otros elementos de la base de datos Oracle.

4. **Visual Basic Code**  
   **Financial Accounting** contiene los componentes del Sistema de Contabilidad Financiera.   
   **Common Library** contiene código de propósito general, específicamente un _Data Access Library_.  
   **Banobras** contiene código específico del Banco como es el caso de los programas que generan los reportes para SIGRO o de la contabilidad fiduciaria.

5. **Web Site**  
   Sitio web donde se ejecuta SICOFIN, basado en páginas de servidor ASP versión 1.0.

# Documentación

Se incluyen los archivos de la garantía, la licencia y las notas de instalación que se entregaron al Banco en su momento. 

El archivo [SCF-SPD-2002-1.pdf](https://github.com/Ontica/Sicofin/blob/master/SCF-SPD-2002-1.pdf) contiene notas generales sobre el diseño de la solución entregada. 

**SCF-SM-2001 Contabilidad financiera.mdl** es un archivo _Rational Rose_ con el modelo estático del sistema (clases).


# Licencia

Este sistema se distribuye bajo la licencia [GNU AFFERO GENERAL PUBLIC LICENSE](https://github.com/Ontica/Sicofin/blob/master/LICENSE.txt).

En Óntica siempre entregamos soluciones y sistemas de información de código abierto. Consideramos que esta práctica es especialmente 
correcta cuando se trata de sistemas de utilidad pública, como es el caso de los sistemas para gobierno.

# Versión

Esta versión del código es la más actualizada, y a nuestro entender, es la que corresponde a la que se entregó oficialmente al Banco en noviembre de 2002.

Publicamos el código del SICOFIN en virtud de que Banobras lo ha mantenido funcionando desde que entró en operación el 5 de febrero de 2001.


# Copyright

Copyright © 1999-2002. La Vía Óntica S.C. y colaboradores.

