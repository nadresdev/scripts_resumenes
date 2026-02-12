
# Diccionario de comentarios para las columnas de las hojas generadas
# Este script ser importado posteriormente para aplicar comentarios a los Excels.

COL_COMMENTS = {
    # Hoja: Resumen_Diario
    "Resumen_Diario": {
        "fecha": "Fecha de creacin del lead [fxCreated]. Agrupacin diaria.",
        "leads_insertados": "Total de leads ingresados en el da.",
        "contactados": "Cantidad de leads con al menos una interaccin (TMO > 0).",
        "ventas": "Cantidad de leads marcados como venta [lastOcmCoding = VENTA / POLIZA].",
        "interacciones_dia": "Suma total de intentos de llamada realizados en el da.",
        "tiempo_total_llamadas_hms": "Tiempo total acumulado en llamadas (HH:MM:SS) sumando todos los intentos.",
        "contactabilidad_%": "Porcentaje de leads contactados sobre insertados (contactados / leads_insertados * 100).",
        "conversion_%": "Porcentaje de ventas sobre leads insertados (ventas / leads_insertados * 100).",
        "mediana_tmo_x_dia_seg": "Mediana del TMO total por registro (en segundos).",
        "mediana_tmo_x_dia_hms": "Mediana del TMO total por registro (formato HH:MM:SS).",
        "mediana_tmo_venta_x_dia_seg": "Mediana del TMO acumulado hasta la venta (en segundos).",
        "mediana_tmo_venta_x_dia_hms": "Mediana del TMO acumulado hasta la venta (formato HH:MM:SS).",
        "mediana_sla_seg": "Mediana del tiempo transcurrido desde ingreso hasta primera llamada (en segundos).",
        "mediana_sla_hms": "Mediana del SLA (formato HH:MM:SS)."
    },
    
    # Hoja: Detalle_Leads_Unicos (Placeholder para cuando desees llenarlo)
    "Detalle_Leads_Unicos": {
        "_id": "Identificador nico del lead.",
        "fullname": "Nombre completo del cliente.",
        # ... puedes agregar ms aqu ...
        "tiempo_total_llamadas_hms": "Suma de tiempos de todas las llamadas (timeCall1...10).",
        "tmo_total_registro": "Suma de TMOs mayores a 0.",
        "tmo_venta": "Suma de TMOs de interacciones no exitosas previas a la venta.",
        "sla_seg": "Tiempo en segundos desde fxCreated hasta fxFirstcall."
    }
}
