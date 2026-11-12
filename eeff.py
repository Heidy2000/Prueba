"""
Estados financieros (balance y PyG) a partir del balance de SAP (transacción F.01), mes a mes.

El archivo de SAP trae, por posición de balance/PyG y por cuenta: Sociedad, Pos. balance/PyG, Texto (cuenta +
nombre), División, Tipo de moneda, Moneda, saldo INICIAL (cierre del mes anterior), saldo FINAL (cierre del mes),
desviaciones. Las filas con Sociedad son cuentas; las demás son subtotales de la jerarquía de SAP.

Las reglas replican las fórmulas del libro 'Balance - indicadores' (hojas ESF Mensual, Estado de resultados Mensual
y EBITDA): cada línea suma el saldo FINAL de las filas que cumplen un criterio:
  C = posición de balance igual al código (solo cuentas si se indica), N = primeros 6 dígitos de la cuenta,
  R = cuenta de 10 dígitos.
El PyG de SAP es acumulado del año (INICIAL de enero = 0); el del mes es la diferencia FINAL − INICIAL.
La porción no corriente (hoja 'No Corriente' del libro, manual) no viene en SAP: se toma la del mes anterior y se
puede ajustar en Administración.
"""

import pandas as pd

# ------------------------------------------------------------------ balance (estado de situación financiera)
# (grupo, subgrupo, línea, regla...)  regla: ('C'|'N'|'R', código, solo_cuentas, signo, índice de su no corriente)
#                                            ('nc',) porción no corriente · ('resultado',) · ('cero',)
ESF = [
    ('ACTIVO CORRIENTE', 'EFECTIVO Y EQUIVALENTES AL EFECTIVO', 'CAJA', 'C', '110510', True, 1, None),
    ('ACTIVO CORRIENTE', 'EFECTIVO Y EQUIVALENTES AL EFECTIVO', 'BANCOS', 'C', '111005', True, 1, None),
    ('ACTIVO CORRIENTE', 'EFECTIVO Y EQUIVALENTES AL EFECTIVO', 'CTS DE AHORROS', 'C', '112005', True, 1, None),
    ('ACTIVO CORRIENTE', 'EFECTIVO Y EQUIVALENTES AL EFECTIVO', 'FONDOS', 'C', '112515', True, 1, None),
    ('ACTIVO CORRIENTE', 'EFECTIVO Y EQUIVALENTES AL EFECTIVO', 'EFECTIVO MANTENIDO EN D. FIDUCIA Y CCA', 'C', '124505', True, 1, None),
    ('ACTIVO CORRIENTE', 'INVERSIONES', 'INVERSIONES', 'C', '120530', True, 1, 30),
    ('ACTIVO CORRIENTE', 'INVERSIONES', 'DETERIORO INVERSIONES', 'C', '', False, -1, 31),
    ('ACTIVO CORRIENTE', 'DERECHO DE USO', 'PROPIEDAD PLANTA Y EQUIPO DERECHOS DE USO', 'C', '159005', True, 1, None),
    ('ACTIVO CORRIENTE', 'DERECHO DE USO', 'PROPIEDAD POR DERECHOS DE USO', 'C', '159270', True, 1, None),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'VENTA SERVICIOS', 'C', '1407', False, 1, 39),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'CLIENTES', 'C', '1305', False, 1, 40),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'DEUDORES VARIOS', 'C', '1380', False, 1, 41),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'ANTICIPOS Y AVANCES', 'C', '1330', False, 1, 42),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'RECLAMACIONES', 'C', '1360', False, 1, 43),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'VALORACION DERIVADOS', 'C', '1395', False, 1, 44),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'PROVISIONES CLIENTES', 'C', '1399', False, 1, 45),
    ('ACTIVO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES', 'CTAS POR COBRAR A VINCULADOS ECONOMICOS', 'C', '1320', False, 1, 46),
    ('ACTIVO CORRIENTE', 'INVENTARIOS CORRIENTES', 'MATERIALES PRESTACION DEL SERVICIO', 'C', '145505', True, 1, None),
    ('ACTIVO CORRIENTE', 'INVENTARIOS CORRIENTES', 'DETERIORO PARA PROTECCION INVENTARIOS', 'C', '1499', False, -1, None),
    ('ACTIVO CORRIENTE', 'ACTIVOS POR IMPUESTOS CORRIENTES', 'ANTICIPO DE IMP. Y CONTRIB. O SALD', 'C', '1355', False, 1, 48),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'TERRENOS', 'C', '150405', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'INSTALACIONES TECNICAS', 'C', '152005', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'MUEBLES Y ENSERES', 'C', '152405', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'MAQUINARIA Y EQUIPO', 'C', '152010', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'PROPIEDADES EN TRANSITO', 'C', '158805', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'EQUIPOS DE COMUNICACION Y COMPUTACION', 'C', '152805', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'EQUIPO DE TRANS, TRACCION Y ELEVACION', 'C', '154005', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'DERECHO DE USO', 'cero'),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'PROVISION MAQUINARIA Y EQUIPO', 'C', '1599', False, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDADES PLANTA Y EQUIPO', 'DEPRECIACION DE PPYE ACUMULADA', 'C', '1592', False, 1, 8),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES', 'INVERSIONES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES', 'DETERIORO INVERSIONES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES EN SUBSIDIARIAS', 'INVERSIONES EN SUBSIDIARIAS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES EN SUBSIDIARIAS', 'CUENTA PUENTE INVERSIONES EN ACCIONES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES EN SUBSIDIARIAS', 'DETERIORO INVERSIONES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES EN ASOCIADAS', 'INVERSIONES EN ASOCIADAS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'INVERSIONES EN NEGOCIOS CONJUNTOS', 'INVERSIONES EN NEGOCIOS CONJUNTOS', 'C', '', False, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDAD DE INVERSIÓN', 'PROPIEDAD DE INVERSIÓN', 'C', '', False, 1, None),
    ('ACTIVO NO CORRIENTE', 'PROPIEDAD DE INVERSIÓN AL COSTO', 'PROPIEDAD DE INVERSIÓN AL COSTO', 'C', '', False, 1, None),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'VENTA SERVICIOS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'CLIENTES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'DEUDORES VARIOS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'ANTICIPOS Y AVANCES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'RECLAMACIONES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'VALORACION DERIVADOS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'PROVISIONES CLIENTES', 'nc'),
    ('ACTIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR NO CORRIENTES', 'CTAS POR COBRAR A VINCULADOS ECONOMICOS', 'nc'),
    ('ACTIVO NO CORRIENTE', 'OTROS ACTIVOS', 'OTROS ACTIVOS', 'C', '170520', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'ACTIVOS POR IMPUESTOS CORRIENTES, NO CORRIENTE', 'ANTICIPO DE IMP. Y CONTRIB. O SALD', 'nc'),
    ('ACTIVO NO CORRIENTE', 'ACTIVO POR IMPUESTO DIFERIDO', 'ACTIVO POR IMPUESTO DIFERIDO', 'R', '1710760001', False, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'SERVICIOS', 'C', '170540', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'IMPUESTOS', 'C', '170555', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'OTROS DIFERIDOS', 'C', '170595', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'DIVERSOS', 'C', '171004', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'HONORARIOS', 'C', '170510', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'PROGRAMAS PARA COMPUTADORES (SOFTWARE)', 'C', '171016', True, 1, None),
    ('ACTIVO NO CORRIENTE', 'INTANGIBLES', 'INTANGIBLES', 'C', '16', False, 1, None),
    ('PASIVO CORRIENTE', 'OBLIGACIONES FINANCIERAS', 'PASIVOS FINANCIEROS A COSTO AMORTIZADO', 'C', '2110', True, -1, 92),
    ('PASIVO CORRIENTE', 'OBLIGACIONES FINANCIERAS', 'CREDITO BANCO SANTANDER', 'N', '210510', False, -1, 93),
    ('PASIVO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'PARTICULARES', 'N', '219505', False, -1, 94),
    ('PASIVO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'COMPAÑIAS VINCULADAS', 'N', '219510', False, -1, 95),
    ('PASIVO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'SOCIOS O ACCIONISTAS', 'N', '219520', False, -1, 96),
    ('PASIVO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'CASA MATRIZ', 'N', '219515', False, -1, 97),
    ('PASIVO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'C/C EMPRESAS GRUPO', 'N', '231505', False, -1, 98),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'DE MATERIALES Y MERCANCÍAS', 'C', '220515', True, -1, 99),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'DE ACTIVOS FIJOS', 'C', '220520', True, -1, 100),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PROVEEDORES DEL EXTERIOR', 'C', '221005', True, -1, 101),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'GASTOS FINANCIEROS', 'C', '233505', True, -1, 102),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'GASTOS LEGALES', 'C', '233510', True, -1, 102),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'LIBROS,SUSCRIPCIONES,PERIODICOS', 'C', '233515', True, -1, 103),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'COMISIONES', 'C', '233520', True, -1, 103),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'HONORARIOS', 'C', '233525', True, -1, 104),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS TECNICOS', 'C', '233530', True, -1, 105),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS DE MANTENIMIENTO', 'C', '233535', True, -1, 106),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'ARRENDAMIENTOS', 'C', '233540', True, -1, 107),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'TRANSPORTES FLETES Y ACARREOS', 'C', '233545', True, -1, 108),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS PUBLICOS', 'C', '233550', True, -1, 109),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SEGUROS', 'C', '233555', True, -1, 110),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PROVEEDORES GASTOS DE VIAJE', 'C', '233560', True, -1, None),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'GASTOS DE REPRESENTACION Y RELACIONES PUBLICAS', 'C', '233565', True, -1, None),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS ADUANEROS', 'C', '233570', True, -1, 112),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'OTROS', 'C', '233595', True, -1, 112),
    ('PASIVO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PASIVOS ESTIMADOS Y PROVISIONES', 'C', '26', False, -1, None),
    ('PASIVO CORRIENTE', 'PASIVOS POR IMPUESTOS CORRIENTES', 'IMPUESTO DE RENTA', 'C', '24', False, -1, 115),
    ('PASIVO CORRIENTE', 'PASIVOS POR IMPUESTOS CORRIENTES', 'RETENCION EN LA FUENTE', 'C', '2365', False, -1, 116),
    ('PASIVO CORRIENTE', 'PASIVOS POR IMPUESTOS CORRIENTES', 'IMPUESTO A LAS VENTAS', 'C', '2367', False, -1, 117),
    ('PASIVO CORRIENTE', 'PASIVOS POR IMPUESTOS CORRIENTES', 'INDUSTRIA Y COMERCIO', 'C', '2368', False, -1, 118),
    ('PASIVO CORRIENTE', 'BENEFICIOS A EMPLEADOS', 'OBLIGACIONES LABORALES', 'C', '25', False, -1, None),
    ('PASIVO CORRIENTE', 'OTROS PASIVOS', 'PASIVO FINANCIERO DERECHO DE USO', 'N', '219595', False, -1, 120),
    ('PASIVO CORRIENTE', 'OTROS PASIVOS', 'RETENCIONES Y APORTES DE NOMINA', 'C', '2370', False, -1, 121),
    ('PASIVO CORRIENTE', 'OTROS PASIVOS', 'ACREEDORES VARIOS', 'C', '2380', False, -1, None),
    ('PASIVO CORRIENTE', 'OTROS PASIVOS', 'OTROS PASIVOS', 'C', '28', False, -1, None),
    ('PASIVO NO CORRIENTE', 'OBLIGACIONES FINANCIERAS', 'PASIVOS FINANCIEROS A COSTO AMORTIZADO', 'nc'),
    ('PASIVO NO CORRIENTE', 'OBLIGACIONES FINANCIERAS', 'CREDITO BANCO SANTANDER', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'PARTICULARES', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'COMPAÑIAS VINCULADAS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'SOCIOS O ACCIONISTAS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'CASA MATRIZ', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS POR PAGAR PARTES RELACIONADAS', 'C/C EMPRESAS GRUPO', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'DE MATERIALES Y MERCANCÍAS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'DE ACTIVOS FIJOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PROVEEDORES DEL EXTERIOR', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'GASTOS LEGALES', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'COMISIONES', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'HONORARIOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS TECNICOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS DE MANTENIMIENTO', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'ARRENDAMIENTOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'TRANSPORTES FLETES Y ACARREOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SERVICIOS PUBLICOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'SEGUROS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PROVEEDORES GASTOS DE VIAJE', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'GASTOS DE REPRESENTACION Y RELACIONES PUBLICAS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'OTROS', 'nc'),
    ('PASIVO NO CORRIENTE', 'CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES', 'PASIVOS ESTIMADOS Y PROVISIONES', 'nc'),
    ('PASIVO NO CORRIENTE', 'PASIVOS POR IMPUESTOS NO CORRIENTES', 'IMPUESTO DE RENTA', 'nc'),
    ('PASIVO NO CORRIENTE', 'PASIVOS POR IMPUESTOS NO CORRIENTES', 'RETENCION EN LA FUENTE', 'nc'),
    ('PASIVO NO CORRIENTE', 'PASIVOS POR IMPUESTOS NO CORRIENTES', 'IMPUESTO A LAS VENTAS', 'nc'),
    ('PASIVO NO CORRIENTE', 'PASIVOS POR IMPUESTOS NO CORRIENTES', 'INDUSTRIA Y COMERCIO', 'nc'),
    ('PASIVO NO CORRIENTE', 'PASIVO POR IMPUESTO DIFERIDO', 'PASIVO POR IMPUESTO DIFERIDO', 'C', '27', False, -1, None),
    ('PASIVO NO CORRIENTE', 'OTROS PASIVOS', 'PASIVO FINANCIERO DERECHO DE USO', 'nc'),
    ('PASIVO NO CORRIENTE', 'OTROS PASIVOS', 'RETENCIONES Y APORTES DE NOMINA', 'nc'),
    ('PASIVO NO CORRIENTE', 'OTROS PASIVOS', 'ACREEDORES VARIOS', 'nc'),
    ('PASIVO NO CORRIENTE', 'OTROS PASIVOS', 'OTROS PASIVOS', 'nc'),
    ('PATRIMONIO', 'CAPITAL EMITIDO', 'CAPITAL EMITIDO', 'C', '31', False, -1, None),
    ('PATRIMONIO', 'GANANCIAS ACUMULADAS', 'GANANCIAS ACUMULADAS', 'C', '3705', False, -1, None),
    ('PATRIMONIO', 'SUPERAVIT POR DONACION', 'SUPERAVIT POR DONACION', 'C', '', False, 1, None),
    ('PATRIMONIO', 'SUPERAVIT NO REALIZADO POR REVALUACIÓN D', 'SUPERAVIT NO REALIZADO POR REVALUACIÓN D', 'C', '', False, 1, None),
    ('PATRIMONIO', 'RESULTADOS DEL EJERCICIO', 'RESULTADOS DEL EJERCICIO', 'resultado'),
    ('PATRIMONIO', 'PRIMA EN COLOCACIÓN DE ACCIONES', 'PRIMA EN COLOCACIÓN DE ACCIONES', 'C', '3205', False, -1, None),
    ('PATRIMONIO', 'OTRAS RESERVAS', 'OTRAS RESERVAS', 'C', '', False, 1, None),
    ('PATRIMONIO', 'ORI', 'ORI', 'C', '3815', False, -1, None),
]
GRUPOS_ESF = ["ACTIVO CORRIENTE", "ACTIVO NO CORRIENTE", "PASIVO CORRIENTE", "PASIVO NO CORRIENTE", "PATRIMONIO"]

# ------------------------------------------------------------------ estado de resultados (formato EBITDA)
# (grupo, línea, criterio, código, solo_cuentas, signo)  signo +1 = ingreso (saldo SAP crédito), −1 = costo/gasto
_R4180 = [("RECOLECCIÓN DOMICILIARIA", "4180100010"), ("GRANDES GENERADORES", "4180100011"),
          ("LAVADO DE AREAS PUBLICAS", "4180100012"), ("RECOGIDA ESCOMBROS", "4180100013"), ("BARRIDO", "4180100014"),
          ("CORTE Y PODA DE ARBOLES", "4180100015"), ("OTROS ESPECIALES", "4180100016"),
          ("INGRESO ESTIMADO SERVICIO DE ASEO", "4180100017"), ("INGRESO POR SUMINISTRO DE CESTAS", "4180100021"),
          ("INGRESOS ESTIMADOS OTRAS ADICIONES", "4180100022"), ("INGRESO INSTALACIÓN DE CESTAS", "4180100023")]
PYG = (
    [("INGRESOS", n, "R", c, False, 1) for n, c in _R4180]
    + [("COSTOS", n, "N", c, False, -1) for n, c in [
        ("COSTOS DE PERSONAL", "720505"), ("HONORARIOS", "731010"), ("ARRENDAMIENTOS", "731020"),
        ("GENERALES", "731035"), ("TRAMITES Y LICENCIAS", "731040"), ("MANTENIMIENTOS Y REPARACIONES", "731045"),
        ("ALOJAMIENTO Y MANUTENCION", "731055"), ("SERVICIOS MÉDICOS, DEPORTIVOS Y CAPTACIÓN", "731065"),
        ("OTROS", "731095"), ("IMPUESTOS Y TASAS", "731515"), ("SEGUROS", ""), ("REPARACIONES LOCATIVAS", "731550"),
        ("OTROS ", "739570"), ("DIVERSOS DE PERSONAL", "720595")]]
    + [("GASTOS DE ADMINISTRACIÓN", n, "C", c, False, -1) for n, c in [
        ("GASTOS DE PERSONAL", "5105"), ("HONORARIOS", "5110"), ("IMPUESTO AL CONSUMO", "5115"), ("ARRENDAMIENTOS", "5120"),
        ("CONTRIBUCIONES Y AFILIACIONES", "5125"), ("SERVICIOS", "5135"), ("GASTOS LEGALES", "5140"),
        ("MANTENIMIENTO Y REPARACIONES", "5145"), ("ADECUACION E INSTALACION", "5150"), ("GASTOS DE VIAJE", "5155"),
        ("DIVERSOS", "5195"), ("SEGUROS", "5130")]]
    + [("DETERIORO", "SEGUROS", "N", "731530", False, -1), ("DETERIORO", "PROVISIONES", "C", "5199", False, -1),
       ("DEPRECIACIONES", "DEPRECIACIÓN ADMINISTRACIÓN", "C", "5160", False, -1),
       ("DEPRECIACIONES", "DEPRECIACIÓN PROPIEDAD PLANTA Y EQUIPO", "N", "732005", False, -1),
       ("AMORTIZACIONES", "AMORTIZACIONES", "C", "5165", False, -1),
       ("AMORTIZACIONES", "AMORTIZACIÓN INTANGIBLES", "N", "731590", False, -1)]
    + [("OTROS INGRESOS", n, "C", c, True, 1) for n, c in [
        ("VENTA DE MATERIALES", "420515"), ("REINTEGRO DE OTROS COSTOS Y GASTOS", "425050"),
        ("APROVECHAMIENTOS", "429505"), ("AJUSTE AL PESO", "429581"),
        ("INGRESO POR VENTA DE INMOVILIZADO MATERIAL", "424595"), ("DESCUENTOS COMERCIALES", "421040"),
        ("PRESTACION DE SERVICIOS", "423595"), ("INGRESOS POR ARRENDAMIENTO MAQUINARIA Y EQUIPO", "422015")]]
    + [("OTROS GASTOS", n, "C", c, False, -1) for n, c in [
        ("PÉRDIDA EN VENTA Y RETIRO DE BIENES", "5310"), ("GASTOS EXTRAORDINARIOS", "5315"),
        ("GASTOS DIVERSOS", "5395")]]
    + [("INGRESOS FINANCIEROS", n, "R", c, False, 1) for n, c in [
        ("INTERESES BANCARIOS", "4210050010"), ("INTERESES DE MORA", "4210050020"),
        ("DIFERENCIAS POSITIVAS DE CAMBIO NO REALIZADAS", "4210200010"),
        ("DIFERENCIAS POSITIVAS DE CAMBIO REALIZADAS", "4210200020"), ("OTROS INGRESOS FINANCIEROS", "4210950010"),
        ("Bº VALORACIÓN INST. FINANCIEROS VALOR RAZONABLE", "4210950020"),
        ("OTROS INGRESOS DERIVADOS", "4210950012")]]
    + [("COSTOS FINANCIEROS", n, "R", c, False, -1) for n, c in [
        ("GASTOS BANCARIOS 4XMIL", "5305050000"), ("COMISIONES BANCARIAS", "5305150000"),
        ("CORRIENTES SOBRE OBLIGACIONES FINANCIERAS", "5305200020"), ("MORA SOBRE OBLIGACIONES", "5305200030"),
        ("INTERESES LEASING USO INMOVILIZADO", "5305200050"),
        ("DIFERENCIA NEGATIVA EN CAMBIO NO REALIZADA", "5305250010"),
        ("DIFERENCIA NEGATIVA EN CAMBIO REALIZADA", "5305250020"),
        ("DESCUENTOS COMERCIALES CONDICIONADOS", "5305350000"), ("CHEQUERAS", "5305550000"),
        ("OTROS GASTOS FINANCIEROS", "5305950000"),
        ("Pª VALORACIÓN INST. FINANCIEROS VALOR RAZONABLE", "5305950010")]]
    + [("IMPUESTO CORRIENTE", "IMPUESTO DE RENTA Y COMPLEMENTARIOS", "R", "5405050000", False, -1),
       ("IMPUESTO DIFERIDO", "IMPUESTO A LAS GANANCIAS DIFERIDO", "R", "5405100000", False, -1)]
)
GRUPOS_PYG = ["INGRESOS", "COSTOS", "GASTOS DE ADMINISTRACIÓN", "DETERIORO", "DEPRECIACIONES", "AMORTIZACIONES",
              "OTROS INGRESOS", "OTROS GASTOS", "INGRESOS FINANCIEROS", "COSTOS FINANCIEROS", "IMPUESTO CORRIENTE",
              "IMPUESTO DIFERIDO"]
SIGNO_GRUPO = {g: (1 if g in ("INGRESOS", "OTROS INGRESOS", "INGRESOS FINANCIEROS") else -1) for g in GRUPOS_PYG}


def _txt(x) -> str:
    return "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()


# ------------------------------------------------------------------ lectura de archivos
def leer_sap(archivo) -> pd.DataFrame:
    """Balance F.01 de SAP: sociedad, posicion, texto, inicial, final (una fila por posición/cuenta)."""
    crudo = pd.read_excel(archivo, header=None, dtype=object)
    fila = next((i for i in range(min(40, len(crudo)))
                 if any("pos. balance" in _txt(v).lower() or "pos.balance" in _txt(v).lower() for v in crudo.iloc[i])),
                None)
    if fila is None:
        raise ValueError("No encuentro el encabezado del balance de SAP ('Pos. balance/PyG').")
    enc = [_txt(v).lower() for v in crudo.iloc[fila]]
    c_soc = next(j for j, t in enumerate(enc) if t.startswith("sociedad"))
    c_pos = next(j for j, t in enumerate(enc) if "pos. balance" in t or "pos.balance" in t)
    c_txt = next((j for j, t in enumerate(enc) if t.startswith("texto")), c_pos + 1)
    c_mon = next((j for j, t in enumerate(enc) if t == "moneda"), None)
    if "inicial" in enc and "final" in enc:
        c_ini, c_fin = enc.index("inicial"), enc.index("final")
    elif c_mon is not None:           # descarga directa: los dos saldos vienen justo después de 'Moneda'
        c_ini, c_fin = c_mon + 1, c_mon + 2
    else:
        raise ValueError("No encuentro las columnas de saldo inicial y final.")
    d = crudo.iloc[fila + 1:]
    out = pd.DataFrame({"sociedad": d.iloc[:, c_soc].map(_txt), "posicion": d.iloc[:, c_pos].map(_txt),
                        "texto": d.iloc[:, c_txt].map(_txt),
                        "inicial": pd.to_numeric(d.iloc[:, c_ini], errors="coerce").fillna(0.0),
                        "final": pd.to_numeric(d.iloc[:, c_fin], errors="coerce").fillna(0.0)})
    out = out[out["posicion"] != ""]
    out["posicion"] = out["posicion"].str.replace(r"\.0$", "", regex=True)
    return out.reset_index(drop=True)


def leer_historico(archivo) -> pd.DataFrame:
    """Archivo de histórico (hojas ESF y PYG con SUBGRUPO 1/2/3 y un mes por columna) ->
    estado, idx, grupo, subgrupo, linea, anio, mes, valor (PyG acumulado del año, en pesos)."""
    filas = []
    libro = pd.read_excel(archivo, sheet_name=None, header=None)
    for hoja, df in libro.items():
        h = hoja.strip().upper()
        if h not in ("ESF", "PYG"):
            continue
        enc = list(df.iloc[0])
        cols_f = [(j, pd.Timestamp(v)) for j, v in enumerate(enc) if isinstance(v, (pd.Timestamp,)) or
                  hasattr(v, "year") and not isinstance(v, (int, float, str))]
        n_txt = 3 if h == "ESF" else 2
        for idx, (_, r) in enumerate(df.iloc[1:].iterrows()):
            if not _txt(r.iloc[0]):
                continue
            g, s = _txt(r.iloc[0]), _txt(r.iloc[1])
            lin = _txt(r.iloc[2]) if n_txt == 3 else s
            for j, f in cols_f:
                v = pd.to_numeric(r.iloc[j], errors="coerce")
                filas.append((h, idx, g, s if n_txt == 3 else g, lin, f.year, f.month, 0.0 if pd.isna(v) else float(v)))
    return pd.DataFrame(filas, columns=["estado", "idx", "grupo", "subgrupo", "linea", "anio", "mes", "valor"])


# ------------------------------------------------------------------ cálculo desde SAP
def _sumar(sap: pd.DataFrame, campo: str, codigo: str, solo_cuentas: bool, col: str = "final") -> float:
    if not codigo:
        return 0.0
    es_cta = sap["sociedad"] != ""
    if campo == "C":
        m = sap["posicion"] == codigo
        if solo_cuentas:
            m &= es_cta
    elif campo == "N":
        m = es_cta & (sap["texto"].str[:6] == codigo)
    else:  # R
        m = es_cta & (sap["texto"].str[:10] == codigo)
    return float(sap.loc[m, col].sum())


def calcular(sap: pd.DataFrame, nc_anterior: dict | None = None) -> pd.DataFrame:
    """Líneas del balance y del PyG (acumulado del año) de un mes desde el balance de SAP.
    nc_anterior: {idx: valor} de las líneas no corrientes del mes anterior (se mantienen)."""
    nc_anterior = nc_anterior or {}
    filas = []
    pyg_vals = []
    for i, (g, ln, campo, cod, solo, sig) in enumerate(PYG):
        v = -sig * _sumar(sap, campo, cod, solo)          # ingreso: −saldo; costo: +saldo (positivos)
        pyg_vals.append(v)
        filas.append(("PYG", i, g, g, ln, v))
    res = sum(SIGNO_GRUPO[PYG[i][0]] * v for i, v in enumerate(pyg_vals))
    esf_vals = {}
    for i, regla in enumerate(ESF):
        tipo = regla[3]
        if tipo == "nc":
            esf_vals[i] = float(nc_anterior.get(i, 0.0))
    for i, regla in enumerate(ESF):
        g, s, ln, tipo = regla[:4]
        if tipo in ("C", "N", "R"):
            _, _, _, campo, cod, solo, sig, resta = regla
            v = sig * _sumar(sap, campo, cod, solo) - (esf_vals.get(resta, 0.0) if resta is not None else 0.0)
        elif tipo == "resultado":
            v = res
        elif tipo == "nc":
            v = esf_vals[i]
        else:
            v = 0.0
        esf_vals[i] = v
        filas.append(("ESF", i, g, s, ln, v))
    return pd.DataFrame(filas, columns=["estado", "idx", "grupo", "subgrupo", "linea", "valor"])


def totales_esf(esf: pd.DataFrame) -> dict:
    """Totales del balance de un mes (DataFrame con grupo y valor)."""
    t = esf.groupby("grupo")["valor"].sum()
    g = {k: float(t.get(k, 0.0)) for k in GRUPOS_ESF}
    g["ACTIVO"] = g["ACTIVO CORRIENTE"] + g["ACTIVO NO CORRIENTE"]
    g["PASIVO"] = g["PASIVO CORRIENTE"] + g["PASIVO NO CORRIENTE"]
    g["PASIVO + PATRIMONIO"] = g["PASIVO"] + g["PATRIMONIO"]
    return g


# ------------------------------------------------------------------ resumen e indicadores por mes
def _idx(linea: str, grupo: str) -> int:
    return next(i for i, r in enumerate(ESF) if r[2] == linea and r[0] == grupo)


def resumen(lineas: pd.DataFrame, tasa_impuesto: float = 0.35) -> pd.DataFrame:
    """Una fila por (anio, mes) con los totales del balance (cierre del mes), el PyG acumulado del año y del mes,
    y los indicadores financieros. 'lineas': estado, idx, grupo, subgrupo, linea, anio, mes, valor."""
    if lineas.empty:
        return pd.DataFrame()
    esf = lineas[lineas["estado"] == "ESF"]
    pyg = lineas[lineas["estado"] == "PYG"]
    meses = sorted(set(zip(lineas["anio"], lineas["mes"])))
    g_esf = esf.pivot_table(index=["anio", "mes"], columns="grupo", values="valor", aggfunc="sum")
    s_esf = esf.pivot_table(index=["anio", "mes"], columns=["grupo", "subgrupo"], values="valor", aggfunc="sum")
    l_esf = esf.pivot_table(index=["anio", "mes"], columns="idx", values="valor", aggfunc="sum")
    g_pyg = pyg.pivot_table(index=["anio", "mes"], columns="grupo", values="valor", aggfunc="sum")
    filas = []
    for a, m in meses:
        def G(k):
            return float(g_esf.loc[(a, m)].get(k, 0.0)) if (a, m) in g_esf.index else 0.0

        def S(g, s):
            return float(s_esf.loc[(a, m)].get((g, s), 0.0)) if (a, m) in s_esf.index else 0.0

        def L(i):
            return float(l_esf.loc[(a, m)].get(i, 0.0)) if (a, m) in l_esf.index else 0.0

        def P(k):
            return float(g_pyg.loc[(a, m)].get(k, 0.0)) if (a, m) in g_pyg.index else 0.0
        r = {"anio": a, "mes": m}
        r.update(AC=G("ACTIVO CORRIENTE"), ANC=G("ACTIVO NO CORRIENTE"), PC=G("PASIVO CORRIENTE"),
                 PNC=G("PASIVO NO CORRIENTE"), PAT=G("PATRIMONIO"))
        r["ACTIVO"], r["PASIVO"] = r["AC"] + r["ANC"], r["PC"] + r["PNC"]
        r["efectivo"] = S("ACTIVO CORRIENTE", "EFECTIVO Y EQUIVALENTES AL EFECTIVO")
        r["clientes"] = L(_idx("CLIENTES", "ACTIVO CORRIENTE"))
        r["deudores"] = L(_idx("DEUDORES VARIOS", "ACTIVO CORRIENTE"))
        r["inventarios"] = S("ACTIVO CORRIENTE", "INVENTARIOS CORRIENTES")
        cxp = "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR CORRIENTES"
        r["cxp"] = S("PASIVO CORRIENTE", cxp)
        r["proveedores"] = r["cxp"] - L(_idx("PASIVOS ESTIMADOS Y PROVISIONES", "PASIVO CORRIENTE"))
        r["deuda_fin"] = S("PASIVO CORRIENTE", "OBLIGACIONES FINANCIERAS") + \
            S("PASIVO NO CORRIENTE", "OBLIGACIONES FINANCIERAS")
        r["activo_fijo"] = S("ACTIVO NO CORRIENTE", "PROPIEDADES PLANTA Y EQUIPO") + \
            S("ACTIVO NO CORRIENTE", "INTANGIBLES")
        for k, g in [("ing", "INGRESOS"), ("cos", "COSTOS"), ("adm", "GASTOS DE ADMINISTRACIÓN"),
                     ("det", "DETERIORO"), ("dep", "DEPRECIACIONES"), ("amo", "AMORTIZACIONES"),
                     ("oi", "OTROS INGRESOS"), ("og", "OTROS GASTOS"), ("if", "INGRESOS FINANCIEROS"),
                     ("cf", "COSTOS FINANCIEROS"), ("impc", "IMPUESTO CORRIENTE"), ("impd", "IMPUESTO DIFERIDO")]:
            r[k] = P(g)
        filas.append(r)
    t = pd.DataFrame(filas).set_index(["anio", "mes"]).sort_index()
    # resultados acumulados del año
    t["bruto"] = t["ing"] - t["cos"]
    t["ebitda"] = t["bruto"] - t["adm"]
    t["operacional"] = t["ebitda"] - t["det"] - t["dep"] - t["amo"]
    t["rai"] = t["operacional"] + t["oi"] - t["og"] + t["if"] - t["cf"]
    t["neto"] = t["rai"] - t["impc"] - t["impd"]
    # valores del mes = acumulado del mes − acumulado del mes anterior (mismo año)
    flujo = ["ing", "cos", "adm", "det", "dep", "amo", "oi", "og", "if", "cf", "impc", "impd", "bruto", "ebitda",
             "operacional", "rai", "neto"]
    prev = t[flujo].groupby(level=0).shift(1)
    primero = t.index.get_level_values(1) == 1
    mes_v = t[flujo] - prev.where(~pd.Series(primero, index=t.index), 0).fillna(0)
    mes_v[primero] = t.loc[primero, flujo]
    for c in flujo:
        t[f"{c}_m"] = mes_v[c]
    # EBITDA últimos 12 meses (solo con meses consecutivos)
    t["ebitda_12m"] = t["ebitda_m"].rolling(12, min_periods=12).sum()
    t["ing_12m"] = t["ing_m"].rolling(12, min_periods=12).sum()
    div = lambda a, b: (a / b).where(b.abs() > 1e-6)                                  # noqa: E731
    t["m_bruto"], t["m_ebitda"] = div(t["bruto"], t["ing"]), div(t["ebitda"], t["ing"])
    t["m_operacional"], t["m_rai"], t["m_neto"] = div(t["operacional"], t["ing"]), div(t["rai"], t["ing"]), \
        div(t["neto"], t["ing"])
    t["endeudamiento"] = div(t["PASIVO"], t["ACTIVO"])
    t["corto_plazo"] = div(t["PC"], t["PASIVO"])
    t["apalancamiento"] = div(t["PASIVO"], t["PAT"])
    t["gasto_fin_neto"] = t["cf"] - t["if"]
    t["cobertura"] = div(t["ebitda"], t["gasto_fin_neto"])
    t["ktnc"] = t["AC"] - t["PC"]
    t["razon_corriente"] = div(t["AC"], t["PC"])
    t["prueba_acida"] = div(t["AC"] - t["inventarios"], t["PC"])
    t["kto"] = t["clientes"] + t["deudores"] + t["inventarios"]           # capital de trabajo operativo
    t["ktno"] = t["kto"] - t["cxp"]
    t["deuda_neta"] = t["deuda_fin"] - t["efectivo"]
    t["deuda_ebitda"] = div(t["deuda_neta"], t["ebitda_12m"])
    meses_t = pd.Series(t.index.get_level_values(1), index=t.index)
    t["roe"] = div(t["neto"] * 12 / meses_t, t["PAT"])                     # anualizado
    t["roa"] = div(t["operacional"] * (1 - tasa_impuesto) * 12 / meses_t, t["kto"] + t["activo_fijo"])
    # rotaciones (acumuladas del año, como la hoja de indicadores)
    dias = meses_t * 30
    prom_cart = (t["clientes"] + t["clientes"].shift(1).fillna(t["clientes"])) / 2
    prom_prov = (t["proveedores"] + t["proveedores"].shift(1).fillna(t["proveedores"])) / 2
    t["dias_cartera"] = div(dias * prom_cart, t["ing"])
    t["dias_proveedores"] = div(dias * prom_prov, t["cos"] + t["adm"])
    t["ciclo_efectivo"] = t["dias_cartera"] - t["dias_proveedores"]
    return t
