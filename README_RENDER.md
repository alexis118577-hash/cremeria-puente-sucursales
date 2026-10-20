# Puente en vivo de sucursales · Render Free · v1.6.51

Esta versión usa Render únicamente como puente de red para consultar otra sucursal y enviar una orden de reabastecimiento mientras las dos aplicaciones están conectadas. Cada computadora conserva su propia base local. Render no recibe publicaciones periódicas ni guarda inventario, ventas, usuarios, reportes o comandos en archivos o bases de datos.

Las consultas se reenvían en vivo por la conexión saliente abierta por cada sucursal. El reabastecimiento también se envía directamente a la sucursal destino por esa conexión; el inventario, el movimiento, el egreso de caja y la auditoría se guardan solamente en la computadora destino. El comando y su respuesta pasan brevemente por la memoria de Render durante la petición, sin cola ni almacenamiento posterior. Si la sucursal está desconectada, la orden no se guarda para después.

## Preparar el servicio en Render

1. Sube el contenido de esta carpeta a la raíz de un repositorio privado.
2. En Render, crea un Blueprint desde ese repositorio y usa `render.yaml`.
3. Conserva el plan `Free` y espera a que termine el despliegue.
4. En Environment, copia el valor generado para `REMOTE_API_KEY`. Guárdalo como secreto y no lo compartas fuera de las computadoras de sucursal.
5. Copia la URL `https://...onrender.com` del servicio.

## Actualizar un servicio existente

Si la ruta `/api/remote/inventory/restock` responde HTTP 404, el código desplegado en Render está desactualizado. Reemplaza en la raíz del repositorio de GitHub los archivos del paquete `Cremeria_Render_Puente_Sucursales_v1_6_51_FREE.zip`, confirma los cambios en `main` y espera el despliegue automático de Render. Conserva la URL y la variable `REMOTE_API_KEY` existentes. Cuando el despliegue esté activo, abre `/health` en el mismo dominio y confirma que muestra `"version":"1.6.51"`.

El Blueprint no crea discos ni bases de datos y la aplicación puente no inicializa almacenamiento local en Render.

## Configurar cada sucursal

En la aplicación Windows v13.6.64, abre **Configuración de sucursal** y captura:

- Un nombre y un código único para esa sucursal.
- La URL de Render.
- La misma `REMOTE_API_KEY`.
- Activa **Habilitar publicación remota** y **Mantener el puente conectado mientras esta sucursal esté abierta**.

Guarda y pulsa **Conectar ahora**. El estado debe aparecer como `ONLINE`. Repite esto con un código distinto en cada sucursal.

## Consultar desde la sucursal principal

En el administrador principal, registra cada sucursal con su nombre, el mismo código único, la URL del servicio Render, autenticación **Token / API key** y la misma `REMOTE_API_KEY`. Pulsa **Probar conexión** y luego **Conectar** en la sucursal. En la micro sesión, abre **Stock de productos**, captura la cantidad en la columna **Reabastecer** y pulsa el botón correspondiente. La sucursal destino valida que tenga habilitado `writeInventory`. El movimiento aumenta únicamente su inventario: no registra una salida de efectivo ni requiere saldo en caja o costo de adquisición.

Solo se permite el reabastecimiento remoto. Los ajustes, descuentos de stock y cambios arbitrarios de inventario siguen bloqueados. La pantalla principal siempre muestra la acción en una micro sesión conectada; la sucursal destino valida el permiso en el momento de ejecutar la orden. Android sigue conectándose al backend local y no se modifica su APK; Windows 13.6.64 mantiene compatibilidad con Android 13.6.48.

## Cuando una sucursal se apaga

Al cerrar la aplicación o apagar su computadora, termina su conexión saliente. No se pueden ver datos de esa sucursal hasta que vuelva a abrir la aplicación y recupere la conexión. Render Free puede reiniciarse; la aplicación vuelve a conectar automáticamente. El primer acceso después de un periodo inactivo puede tardar mientras Render inicia.

La conexión envía un pequeño mensaje de mantenimiento cada minuto para que Render Free no considere inactiva una sucursal que sigue abierta. Esto no copia datos de negocio.
