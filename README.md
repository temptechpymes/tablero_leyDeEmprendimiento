Este proyecto consiste en un flujo de Power Automate, un script de Python y un tablero de Power Bi.

# tablero_leyDeEmprendimiento

La finalidad del tablero es tener una perspectiva sobre la Ley de emprendimiento en Cesde.
Los indicadores que contiene el tablero son:

- Personas matriculadas.
- Cursos activos.
- Usuarios certificados.
- Usuarios en progreso.
- Usuarios suspendidos.
- Cursos certificados.
- Usuarios activos.
- Porcentaje de avance por año.
- Curso/diplomado y usuarios suspendidos.
- Curso y usuarios certificados.
- Matriz con información detallada de cada curso.

Estos indicadores están regidos por 3 filtros:
Año, Mes y Curso.


##Cálculos

-Cursos Activos = COUNT('P Y F'[curso]):se utiliza para calcular el 
número total de cursos activos en la tabla llamada 'P Y F.

-Cursos Certificados = COUNTROWS(FILTER('P Y F', 'P Y F'[estado_curso] = "Certificado")): se utiliza para contar el número 
de cursos que están marcados como "Certificado" en la columna 
'estado_curso' de la tabla 'P Y F'.

-No Iniciado = COUNTROWS(FILTER('P Y F','P Y F'[estado_curso] = "No iniciado")): se utiliza para contar el número de cursos 
que están marcados como "No iniciado" en la columna 'estado_curso' 
de la tabla 'P Y F'.

-Personas Matriculadas = 'Medidas'[Usuarios Suspendidos] + 'Medidas'[Cursos Activos]: se utiliza para calcular el total de 
personas matriculadas en los cursos que se encuentran activos.

-Recuento de Curso = COUNT(Certificados[Curso]): se utiliza para contar el número de filas en la columna "Curso" de 
la tabla "Certificados".

-Usuarios Activos = CALCULATE(DISTINCTCOUNT('P Y F'[email]), FILTER('P Y F', 'P Y F'[estado_curso] = "Certificado" || 
'P Y F'[estado_curso] = "En progreso")): se utiliza para calcular el número de usuarios activos en función del estado del curso 
en la tabla 'P Y F'.

-Usuarios Certificados = CALCULATE(DISTINCTCOUNT('P Y F'[email]), FILTER('P Y F', 'P Y F'[estado_curso] = "Certificado")): se utiliza 
para calcular el número de usuarios que tienen cursos con estado "Certificado" en la tabla 'P Y F'.

-Usuarios En Progreso = COUNTROWS(FILTER('P Y F', 'P Y F'[estado_curso] = "En progreso")): se utiliza para contar el número de usuarios 
cuyo estado de curso es "En progreso" en la tabla 'P Y F'.

-Usuarios Suspendidos = COUNT('Usuarios suspendidos'[Número de ID]): se utiliza para contar el número de usuarios suspendidos en la tabla 'Usuarios suspendidos'.


FLUJO LEY DE EMPRENDIMIENTO Y SUSCRIPCIÓN EMPRESARIAL

Este flujo se encuentra diseñado en Power Automate.

Con Power Automate, los usuarios pueden crear flujos de trabajo sin necesidad de conocimientos de programación, utilizando una interfaz gráfica intuitiva. Pueden automatizar tareas y procesos repetitivos, como la recolección y procesamiento de datos, la notificación de eventos, la sincronización de datos entre aplicaciones, la aprobación de documentos y mucho más.

Power Automate se integra con una amplia gama de servicios y aplicaciones populares, como Office 365, SharePoint, Dynamics 365, Teams, Salesforce, Twitter, entre otros. Esto permite a los usuarios conectar diferentes sistemas y aprovechar las capacidades de automatización en diversos escenarios empresariales.

También se encuentran los archivos de excel que son la fuente de datos, si el proyecto se retoma, actualizarán la forma de obtenerlos desde la web y las plataformas del CESDE.

Y por último para el proyecto de suscripción empresarial se corre un script de Python que realiza la estructuración de los datos obtenidos. "AutomatizacionCESDE.py"


 









