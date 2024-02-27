### Argentum Online 0.11.5

# Changelog:

- 0.11.5

 * 01/03/2006: Implementé el nuevo inventario gráfico y eliminé el viejo (Maraxus)
 * 01/03/2006: Implementé el nuevo engine de sonido y eliminé todos los sistemas viejos (Maraxus)
 * 06/03/2006: Creacion de constantes y enums varios (mensajes, heading) (AlejoLp)
 * 06/03/2006: Elimine las funciones Move[NSWE] y las converti a MoveTo(E_Heading) (AlejoLp)
 * 06/03/2006: Incorporación de los nuevos managers de surfaces y su interfaz. La implementación queda pendientes hasta estar lsito el nuevo AOSetup (Maraxus)
 * 06/03/2006: Corrección de todos los bugs reportados y cerrados en el bug tracker de SF (Maraxus).
 * 06/03/2006: Incorporación del formulario para cambios del MOTD (OuTIme) con algunos agregados míos (Maraxus).
 * 11/03/2006: Implementación definitiva de los nuevos managers de texturas. (Maraxus).
 * 11/03/2006: Implementación del ClientSetup en vez del viejo y obsoleto RenderMod. (Maraxus).
 * 11/03/2006: Se eliminaron variables viejas o mal tipadas. (Maraxus).
 * 15/03/2006: Se mejoró el sistema anticheats. (Maraxus).
 * 18/03/2006: Corregí varios bugs que impedían que uno se loguease. (Maraxus).
 * 18/03/2006: Corregí un bug que te hacía perder puntos de skills al no actualizar los valores al asignar. (Maraxus).
 * 18/03/2006: Corregí un bug que te permitía asignar un skill por encima de 100. (Maraxus).
 * 30/03/2006: Corregí bugs que evitaban la conexión del cliente con seguridad de Alkon. (Maraxus).
 * 31/03/2006: El NLOGIN no envía losa tributos, ya que estos no son usados por el servidor. (Maraxus).
 * 1/04/2006: Corregí el bug que no actualizaba la posición del rect donde debía renderizar. (Maraxus).
 * 1/04/2006: El ValCode y MD5 se llevaron al final de OLOGIN para evitar problemas si el MD5 tenía comas. (Maraxus).
 * 1/04/2006: Eliminé el head y body del mensaje NLOGIN (eran ceros). (Maraxus).
 * 6/04/2006: Agregué un error manager al insertar elementos a la lista de surfaces dinámica (a algunos OS no le gusta y deben usar el estático). (Maraxus).
 * 6/04/2006: Corregí un bug que permitía utilizar el item de arriba a la izquierda al hacer doble click en un slot vacío. (Maraxus).
 * 6/04/2006: Corregí los accesos inválidos con el AO Dinámico en OS como XP. (Maraxus).
 * 6/04/2006: Corregí un bug en el AO dinámico que producía una pérdida de performance al llenarse la memoria. (Maraxus).
 * 12/04/2006: Corregí numerosos bugs con el engine de sonido, y un problema de performance con el inventario gráfico. (Maraxus).
 * 12/04/2006: Corregí bugs varios a lo largo del código. (Maraxus).
 * 14/04/2006: Los RMs no se escuchan al caminar. (Maraxus).
 * 14/04/2006: Corregí bugs varios. (Maraxus).
 * 14/04/2006: Mejoré el engine de sonido. (Maraxus).
 * 14/04/2006: Al activar la música reproduce el midi correcto para el mapa y no el 2 que estaba por default. (Maraxus).
 * 19/04/2006: La bóveda ahora muestra los datos correctamente. (Maraxus).
 * 19/04/2006: Se corrigió un rt 9 en el inventario gráfico al tener el oro como item elegido. (Maraxus).
 * 19/04/2006: Las coords del char se muestran corerctamente al loguear y al apretar la L. (Maraxus).
 * 20/04/2006: corregí un bug que no desactivaba los sonidos de pasos propios y lluvia al desactivar los sonidos. (Maraxus).
 * 21/04/2006: El panelGM y la lista de guilds no se esconden al agregarse texto a la consola. (Maraxus).

# NOTAS TÉCNICAS:

Es nescesario instalar los siguientes componentes antes de continuar si todavia no cuenta con ellos:
- DirectX (htttp://www.microsoft.com/dx/)
- Visual Basic Runtimes 6 SP 6 (http://support.microsoft.com/default.aspx?scid=kb;en-us;290887)

# Requerimientos mínimos:
- Pentium 233 (o 166 MMX)
- 32 Mb de RAM
- Placa de Video SVGA de 2 Mb compatible con DirectX 7
- Windows 95 o superior
- Mouse, teclado y sonido si pretendes escuchar algo ;)...

# Requerimientos recomendados:
- Pentium III 800 Mhz
- 96 Mb de RAM
- Placa de Video SVGA de 16 Mb compatible con DirectX 7
- Windows XP
- Mouse, teclado y sonido si pretendes escuchar algo ;)...

# Problemas de rendimiento:

- Modos de colores distintos de 16 bits: En algunas ocasiones el juego puede presentar problemas al no ejecutarse en 16 bits de colores. Algunas configuraciónes pueden tener problemas con DirectDraw que hacen que los graficos tomen cantidades exorbitantes de tiempo para ser dibujado. Puede actualizar los drivers de la placa de video o la versión de DirectX. En algunas ocasiones hemos comprobado que esto no soluciona el problema, en esos casos debe modificar la configuración manualmente o activar el cambio de resolución al iniciar.

- Sonido activado en maquinas antiguas: El sonido puede disminuir mucho la performance del juego especialmente si se lo activa en modo 3d en hardware que no lo soporte. Recomendamos desactivarlo en equipos inferiores a Pentium II. De todas formas, puede probar desactivando solamente la función 3D.

- Alphablending y Colored Render: Los efectos de Alphablending y Colored Render son efectuados por software. Si bien esto no deberia representar un problema en la mayoría de los equipos fabricados en los últimos años puede traer problemas con equipos más antiguos, en especial aquellos cuyo CPU no es compatible com MMX. Se recomeinda desactivarlos si su PC fue fabricada antes del 2000.

- Otros problemas: Si tiene otro problema para ejecutar el juego contamos con un foro de ayuda. Puede visitarlo en la sección "Foro de nuestra página" (http://www.argentumonline.com.ar/)

# ANTES DE JUGAR:
Reglamento: Para evitar problemas deberías leer el reglamento del juego, alli se detallan las normas para seguir dentro del juego. Lo puedes encontrar en nuestra página web: http://www.argentumonline.com.ar/.

Manual: Es conveniente leer el manual antes de consultar. Lo puedes encontrar en nuestra página web: http://www.argentum-online.com.ar/.

Foro: Es el medio de consulta más útil donde hay cantidades importantes de otros jugadores dispuestos a ayudarte. Podes ingresar en http://foro.alkon.com.ar/forumdisplay.php?f=98.
