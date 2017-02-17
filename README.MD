# INTRODUCCIÓN

* Descripción: Utilidad que hace click en un botón indicado de una ventada indicada en el fichero HellBoy.config, con la periodicidad indicada.

El ejecutable (HellBoy.exe) y el fichero de configuración (HellBoy.config) deben estar en la misma ruta.

El Fichero de configuración tiene el siguiente aspecto:


>HellBoy v1

>----------

>WindowTitle: Reemplazar

>ButtonCaption: Cancelar

>SleepSeconds: 1

>TTLifeInSecs: 180


Cuando se ejecuta HellBoy.exe la utilidad busca (con la periodicidad indicada en SleepSeconds) una ventana con el nombre indicado en 'WindowTitle', si la encuentra busca en esa ventana un botón con el nombre indicado en 'ButtonCaption' y hace click en él.

La aplicación se ejecuta en bucle con la periodicidad indicada en 'SleepSeconds' durante el número de segundos indicado en 'TTLifeInSecs', si se desea que la aplicación se ejecute indefinidamente dejar 'TTLifeInSecs' sin valor (TTLifeInSecs: ).

La utilidad puede ser llamada desde línea de comandos de la siguiente forma:

``
HellBoy.exe "Reemplazar" "Cancelar" 1 180
``