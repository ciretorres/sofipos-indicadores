"""Validaciones de cobertura y consistencia de los datos del libro."""

from __future__ import annotations

from .be import Boletin
from .config import MESES, MES_NUM
from .tablero import Tablero
from .utilidades import fecha_mes


def _num(x) -> bool:
    """Indica si ``x`` es un entero o un decimal."""
    return isinstance(x, (int, float))


def _tolera_dif(dif: float, publicado: float) -> bool:
    """Indica si una diferencia cabe en el margen de NICAP/IMOR/ICOR.

    Tolera la absoluta de redondeo (0.02 pts) o una relativa al valor
    publicado (1e-5), para no marcar como falla una diferencia grande
    causada por un valor publicado enorme (caso Stori 027032, mayo 2024).
    """
    return dif <= 0.02 or dif <= 1e-5 * abs(publicado)


def validar_datos(
    tablero: Tablero,
    filas: dict[int, list[dict]],
    universo: list[dict],
    be_por_archivo: dict[str, Boletin],
) -> bool:
    """Valida la cobertura y consistencia de los datos del libro.

    Comprueba que cada mes incluya todas las sociedades del universo y
    compara los totales, datos de boletines e indicadores calculados con
    sus valores de referencia.

    Args:
        tablero: Datos originales del tablero, incluidos los valores por
            entidad y los totales del sector.
        filas: Filas preparadas para el libro, agrupadas por año.
        universo: Sociedades que deben aparecer en cada mes.
        be_por_archivo: Boletines disponibles, indexados por nombre de
            archivo.

    Returns:
        ``True`` si todas las validaciones pasan; ``False`` si se detecta
        una diferencia fuera de los márgenes aceptados.
    """
    ok = True
    print("\n=== Validación de datos ===")
    claves = [u["clave"] for u in universo]

    # Comprueba que cada mes tenga una fila por sociedad del universo.
    print("\n  Cobertura por mes:")
    for anio, lista in filas.items():
        por_mes: dict[str, int] = {}
        for f in lista:
            por_mes[f["mes"]] = por_mes.get(f["mes"], 0) + 1

        malos = {
            mes: cantidad
            for mes, cantidad in por_mes.items()
            if cantidad != len(universo)
        }
        print(
            f"    {anio}: {len(lista)} filas · {len(por_mes)} meses · "
            f"{'todas con ' + str(len(universo)) if not malos else malos}"
        )
        if malos:
            ok = False

    # Compara la suma de las sociedades con el total del sector del tablero.
    print(
        "\n  Consistencia del tablero: suma de entidades vs fila 27 "
        "(total del sector):"
    )
    campos = [
        "activos",
        "cartera_e12",
        "cartera_e3",
        "cartera_total_cruda",
        "captacion",
        "capital_neto",
        "requerimiento",
    ]
    etiquetas = {
        "activos": "Activos",
        "cartera_e12": "Cartera E1+E2",
        "cartera_e3": "Cartera E3",
        "cartera_total_cruda": "Cartera total",
        "captacion": "Captación",
        "capital_neto": "Capital Neto",
        "requerimiento": "Requerimiento",
    }

    for campo in campos:
        comparaciones = 0
        peor = 0.0

        # Los meses salen de las filas del libro, no de la config completa.
        for anio, lista in filas.items():
            for mes in sorted({MES_NUM[f["mes"]] for f in lista}):
                fecha = fecha_mes(anio, mes)
                esperado = tablero.sector.get((fecha, campo))
                if esperado is None:
                    continue

                suma = sum(
                    valores.get(campo, 0.0)
                    for (clave, f), valores in tablero.valores.items()
                    if f == fecha
                )
                comparaciones += 1
                peor = max(peor, abs(suma - esperado))

        estado = "OK" if peor <= 1.0 else f"DIF {peor:,.2f}"
        if peor > 1.0:
            ok = False

        print(
            f"    {etiquetas[campo]:<18} {comparaciones:>2} meses · "
            f"peor diferencia {peor:,.2f} pesos · {estado}"
        )

    # Contrasta siete conceptos con los boletines BE de los años del libro.
    pares = [
        ("activos", "activos", "Activos"),
        ("cartera_e12", "cartera_1_2", "Cartera E1+E2"),
        ("cartera_e3", "cartera_3", "Cartera E3"),
        ("estimaciones", "estimaciones", "Estimaciones"),
        ("ingresos_int", "ingresos_int", "Ingresos por intereses"),
        ("gastos_int", "gastos_int", "Gastos por intereses"),
        ("gastos_adm", "gastos_adm", "Gastos de administración"),
    ]
    excepciones_be = {("202606", "027019", "gastos_adm")}

    # Sin el año del boletín en el libro no se cruza (--anios sin 2026).
    boletines = []
    for archivo, boletin in be_por_archivo.items():
        fecha = archivo.split("_")[-1].replace(".xlsx", "")
        anio, mes = int(fecha[:4]), int(fecha[4:])
        if anio not in filas:
            continue
        boletines.append((fecha, anio, mes, boletin))

    if boletines:
        boletines.sort(key=lambda b: (b[1], b[2]))
        nombres_meses: list[str] = []
        for _, _, mes, _ in boletines:
            nombre = MESES[mes].lower()
            if nombre not in nombres_meses:
                nombres_meses.append(nombre)
        anios_validados = {b[1] for b in boletines}
        if len(nombres_meses) == 1:
            lista_meses = nombres_meses[0]
        elif len(nombres_meses) == 2:
            lista_meses = " y ".join(nombres_meses)
        else:
            lista_meses = ", ".join(nombres_meses[:-1]) + " y " + nombres_meses[-1]
        if len(anios_validados) == 1:
            lista_meses = f"{lista_meses} de {next(iter(anios_validados))}"
        print(f"\n  Tablero vs boletín BE ({lista_meses}, miles de pesos):")

    for fecha, anio, mes, boletin in boletines:
        por_clave = {r["clave"]: r for r in boletin.registros}
        filas_mes = {
            f["clave"]: f
            for f in filas[anio]
            if f["mes"] == MESES[mes]
        }
        tipo_serie = (
            "serie reexpresada" if fecha == "202603" else "corte definitivo"
        )
        print(f"    {MESES[mes]} {anio} ({tipo_serie}):")

        for campo_tab, campo_be, etiqueta in pares:
            comparaciones = 0
            peor = 0.0
            peor_clave = None
            discrepancias = []

            for clave, f in filas_mes.items():
                be = por_clave.get(clave)
                if (
                    not be
                    or not _num(f.get(campo_tab))
                    or not _num(be.get(campo_be))
                ):
                    continue

                comparaciones += 1
                dif = abs(f[campo_tab] - be[campo_be])
                if dif > 0.02:
                    discrepancias.append(clave)
                if dif > peor:
                    peor, peor_clave = dif, clave

            if fecha == "202606":
                no_documentadas = [
                    clave
                    for clave in discrepancias
                    if (fecha, clave, campo_tab) not in excepciones_be
                ]
                if no_documentadas:
                    ok = False

                if not discrepancias:
                    estado = "OK"
                elif no_documentadas:
                    estado = f"REVISAR {len(no_documentadas)} entidades"
                else:
                    estado = "OK · excepción documentada 027019 SAE"
            else:
                estado = (
                    "OK"
                    if not discrepancias
                    else f"{len(discrepancias)} diferencias por reexpresión"
                )

            extra = (
                f" · peor {peor:,.4f} ({peor_clave})"
                if peor_clave
                else ""
            )
            print(
                f"      {etiqueta:<24} {comparaciones:>2} entidades"
                f"{extra} · {estado}"
            )

    # Accede a los importes originales en pesos para evitar el redondeo a miles.
    def crudo(f):
        """Obtiene los valores originales del tablero para una fila.

        Busca los datos usando la clave de la sociedad y la fecha de la fila.
        Devuelve un diccionario vacío si no encuentra esa combinación.
        """
        return tablero.valores.get(
            (f["clave"], fecha_mes(f["anio"], MES_NUM[f["mes"]])),
            {},
        )

    # Compara los ROA y ROE calculados con los indicadores publicados.
    # Se admite una diferencia de 0.02 puntos por el redondeo del indicador.
    print("\n  ROA / ROE calculados (RN 12M / promedio 12M) vs publicados:")
    for nombre, num, den, pub in [
        ("ROA", "rn_12m", "activo_prom_12m", "roa_pub"),
        ("ROE", "rn_12m", "capital_prom_12m", "roe_pub"),
    ]:
        comparaciones = 0
        peor = 0.0
        peor_fila = None

        for anio, lista in filas.items():
            for f in lista:
                v = crudo(f)
                if not (
                    _num(v.get(num))
                    and _num(v.get(den))
                    and _num(f.get(pub))
                ):
                    continue
                if v.get(den) == 0:
                    continue

                calculado = v[num] / v[den] * 100
                dif = abs(calculado - f[pub])
                comparaciones += 1
                if dif > peor:
                    peor, peor_fila = dif, f

        estado = "OK" if peor <= 0.02 else f"DIF {peor:,.4f}"
        if peor > 0.02:
            ok = False

        extra = (
            f" ({peor_fila['clave']} {peor_fila['mes']} "
            f"{peor_fila['anio']})"
            if peor_fila
            else ""
        )
        print(
            f"    {nombre}: {comparaciones:>3} comparaciones · "
            f"peor diferencia {peor:,.4f} pts · {estado}{extra}"
        )

    # Contrasta NICAP, IMOR e ICOR calculados con sus valores publicados.
    print("\n  NICAP / IMOR / ICOR calculados vs publicados:")
    pruebas = [
        (
            "NICAP",
            lambda v: v["capital_neto"] / v["requerimiento"] * 100,
            "nicap_pub",
            ("capital_neto", "requerimiento"),
        ),
        (
            "IMOR",
            lambda v: (
                v["cartera_e3"]
                / (v["cartera_e12"] + v["cartera_e3"])
                * 100
            ),
            "imor_pub",
            ("cartera_e3", "cartera_e12"),
        ),
        (
            "ICOR",
            lambda v: abs(v["estimaciones"]) / v["cartera_e3"] * 100,
            "icor_pub",
            ("estimaciones", "cartera_e3"),
        ),
    ]

    for nombre, calculo, campo_pub, requeridos in pruebas:
        comparaciones = 0
        peor = 0.0
        peor_fila = None
        # Peor diferencia que NO pasa la tolerancia (la falla se decide
        # por comparación: una dif grande tolerada no oculta una chica).
        peor_falla = 0.0
        peor_falla_fila = None

        for anio, lista in filas.items():
            for f in lista:
                v = crudo(f)
                if not all(_num(v.get(campo)) for campo in requeridos):
                    continue
                if not _num(f.get(campo_pub)):
                    continue
                if nombre == "NICAP" and not v["requerimiento"]:
                    continue
                if nombre in ("IMOR", "ICOR") and (
                    v["cartera_e3"] == 0
                    or (
                        nombre == "IMOR"
                        and v["cartera_e12"] + v["cartera_e3"] == 0
                    )
                ):
                    continue

                dif = abs(calculo(v) - f[campo_pub])
                comparaciones += 1
                if dif > peor:
                    peor, peor_fila = dif, f
                if not _tolera_dif(dif, f[campo_pub]) and dif > peor_falla:
                    peor_falla, peor_falla_fila = dif, f

        hay_falla = peor_falla_fila is not None
        if hay_falla:
            ok = False
        # Al fallar se reporta la peor fuera de tolerancia; al estar OK,
        # la peor diferencia en general.
        mostrar = peor_falla if hay_falla else peor
        mostrar_fila = peor_falla_fila if hay_falla else peor_fila
        estado = "OK" if not hay_falla else f"DIF {mostrar:,.4f}"

        extra = (
            f" ({mostrar_fila['clave']} {mostrar_fila['mes']} "
            f"{mostrar_fila['anio']})"
            if mostrar_fila
            else ""
        )
        print(
            f"    {nombre}: {comparaciones:>3} comparaciones · "
            f"peor diferencia {peor:,.4f} pts · {estado}{extra}"
        )

    # Comprueba que la cartera total coincida con la suma de E1+E2 y E3.
    print("\n  Cartera total vs boletín (E1+E2+E3):")
    for anio, lista in filas.items():
        comparaciones = 0
        peor = 0.0

        for f in lista:
            campos_cartera = (
                "cartera_e12",
                "cartera_e3",
                "cartera_total_cruda",
            )
            if all(_num(f.get(campo)) for campo in campos_cartera):
                comparaciones += 1
                diferencia = abs(
                    (f["cartera_e12"] + f["cartera_e3"])
                    - f["cartera_total_cruda"]
                )
                peor = max(peor, diferencia)

        estado = "OK" if peor <= 0.02 else "REVISAR"
        print(
            f"    {anio}: {comparaciones} entidades-mes · "
            f"peor diferencia {peor:,.4f} miles · {estado}"
        )
        if peor > 0.02:
            ok = False

    _ = claves
    return ok
