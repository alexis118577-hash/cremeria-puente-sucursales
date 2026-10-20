# Puente en vivo de sucursales · Render Free

Esta versión usa Render únicamente como puente de red para consultar otra sucursal. Cada computadora conserva su propia base local. Render no recibe publicaciones periódicas ni guarda inventario, ventas, usuarios, reportes o comandos en archivos o bases de datos.

Para responder una consulta, el puente recibe una solicitud del administrador, la envía por una conexión saliente abierta por la sucursal y devuelve la respuesta en vivo. Mientras se procesa, la solicitud y su respuesta pasan brevemente por la memoria del servicio. Al terminar la petición no se conserva una copia.

## Preparar el servicio en Render

1. Sube el contenido de esta carpeta a la raíz de un repositorio privado.
2. En Render, crea un Blueprint desde ese repositorio y usa `render.yaml`.
3. Conserva el plan `Free` y espera a que termine el despliegue.
4. En Environment, copia el valor generado para `REMOTE_API_KEY`. Guárdalo como secreto y no lo compartas fuera de las computadoras de sucursal.
5. Copia la URL `https://...onrender.com` del servicio.

El Blueprint no crea discos ni bases de datos y la aplicación puente no inicializa almacenamiento local en Render.

## Configurar cada sucursal

En la aplicación Windows v13.6.56, abre **Configuración de sucursal** y captura:

- Un nombre y un código único para esa sucursal.
- La URL de Render.
- La misma `REMOTE_API_KEY`.
- Activa **Habilitar publicación remota** y **Mantener el puente conectado mientras esta sucursal esté abierta**.

Guarda y pulsa **Conectar ahora**. El estado debe aparecer como `ONLINE`. Repite esto con un código distinto en cada sucursal.

## Consultar desde la sucursal principal

En el administrador principal, registra cada sucursal con su nombre, el mismo código único, la URL del servicio Render, autenticación **Token / API key** y la misma `REMOTE_API_KEY`. Pulsa **Probar conexión** y luego consulta el resumen, inventario, ventas o reportes.

El acceso remoto es de solo lectura. Los cambios de inventario se realizan en la computadora local de la sucursal. Android sigue conectándose al backend local, sin cambios a la versión Android 13.6.49.

## Cuando una sucursal se apaga

Al cerrar la aplicación o apagar su computadora, termina su conexión saliente. No se pueden ver datos de esa sucursal hasta que vuelva a abrir la aplicación y recupere la conexión. Render Free puede reiniciarse; la aplicación vuelve a conectar automáticamente. El primer acceso después de un periodo inactivo puede tardar mientras Render inicia.

La conexión envía un pequeño mensaje de mantenimiento cada minuto para que Render Free no considere inactiva una sucursal que sigue abierta. Esto no copia datos de negocio.
