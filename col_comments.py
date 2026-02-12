
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
        "mediana_sla_hms": "Mediana del SLA (formato HH:MM:SS).",
        "sla_operativo_mediana_hms": "Mediana del SLA para leads creados en horario operativo (L-V 10:00-18:00) [HH:MM:SS].",
        "sla_extra_mediana_hms": "Mediana del SLA para leads creados fuera del horario operativo (L-V resto de horas) [HH:MM:SS].",
        "sla_fds_mediana_hms": "Mediana del SLA para leads creados en fines de semana (Sb-Dom) [HH:MM:SS].",
        "timeAcw_mediana_dia_hms": "Mediana del ACW total por lead (suma de timeAcw1 a 10) para leads contactados [HH:MM:SS]."
    },
    
    # Hoja: Resumen_Agentes
    "Resumen_Agentes": {
        "agente": "Nombre del agente (extrado de callAgent 1-10).",
        "leads_cerrados": "Cantidad leads donde el agente fue el ltimo en gestionar (coincide con lastOcmAgent).",
        "int_contacto": "Cantidad de interacciones (llamadas) con TMO > 0.",
        "int_sin_contacto": "Cantidad de interacciones (llamadas) con TMO = 0.",
        "interacciones_ventas": "Cantidad de interacciones donde el resultado fue VENTA o POLIZA.",
        "interacciones_total": "Suma total de interacciones (contacto + sin contacto).",
        "ventas": "Cantidad de ventas logradas (donde resultDesc de la interaccin es VENTA).",
        "conversion_contactos_%": "Porcentaje de ventas sobre interacciones con contacto (ventas / int_contacto * 100).",
        "conv_contactado_cerrado_%": "Porcentaje de ventas sobre leads cerrados por el agente (ventas / leads_cerrados * 100).",
        "tmo_total_hms": "Suma total del tiempo hablado (TMO) por el agente en todas sus llamadas [HH:MM:SS].",
        "tmo_total_mediana_hms": "Mediana del TMO (> 0) por llamada [HH:MM:SS].",
        "tmo_ventas_hms": "Suma del TMO de las llamadas especficas de venta [HH:MM:SS].",
        "tmo_ventas_mediana_hms": "Mediana del TMO de las llamadas de venta [HH:MM:SS].",
        "tiempo_total_llamadas_hms": "IDEM TMO Total HMS (Suma TMO).",
        "sla_hms_medio": "Promedio del SLA para leads donde el agente hizo la primera llamada.",
        "sla_mediana_hms": "Mediana del SLA para leads donde el agente hizo la primera llamada (callAgent10).",
        "sla_operativo_mediana_hms": "Mediana SLA (Horario Operativo) para leads iniciados por el agente.",
        "sla_extra_mediana_hms": "Mediana SLA (Horario Extra) para leads iniciados por el agente.",
        "sla_fds_mediana_hms": "Mediana SLA (Fines de Semana) para leads iniciados por el agente.",
        "time_acw_hms": "Suma total de tiempos ACW del agente [HH:MM:SS].",
        "acw_mediana_hms": "Mediana del ACW por interaccin del agente [HH:MM:SS].",
        "tmo_no_venta_hms": "Suma TMO de llamadas que NO fueron venta.",
        "tmo_no_venta_mediana_hms": "Mediana TMO de llamadas que NO fueron venta."
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
