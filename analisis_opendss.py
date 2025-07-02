import win32com.client
import math

# Iniciar OpenDSS
dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
if not dssObj.Start(0):
    raise RuntimeError("No se pudo iniciar OpenDSS Engine")

dssText = dssObj.Text
dssCircuit = dssObj.ActiveCircuit

def calcular_thd_bus(bus_name):
    """Funci\u00f3n para calcular THD de un bus espec\u00edfico"""
    dssCircuit.SetActiveBus(bus_name)
    num_nodos = dssCircuit.ActiveBus.NumNodes
    
    if num_nodos == 0:
        return 0.0
    
    # Obtener voltaje fundamental
    dssText.Command = 'Set harmonic=1'
    dssText.Command = 'Solve'
    dssCircuit.SetActiveBus(bus_name)
    v_fund = list(dssCircuit.ActiveBus.VMagAngle)[::2][0] if dssCircuit.ActiveBus.VMagAngle else 0
    
    # Calcular arm\u00f3nicos
    suma_cuad_armonicos = 0.0
    for h in range(2, 16):
        dssText.Command = f'Set harmonic={h}'
        dssText.Command = 'Solve'
        dssCircuit.SetActiveBus(bus_name)
        voltajes_h = list(dssCircuit.ActiveBus.VMagAngle)[::2]
        if voltajes_h:
            suma_cuad_armonicos += voltajes_h[0]**2
    
    return 100 * math.sqrt(suma_cuad_armonicos) / v_fund if v_fund > 0 else 0.0

def obtener_flujos_linea(line_name):
    """Funci\u00f3n para obtener flujos de potencia de una l\u00ednea"""
    try:
        dssCircuit.Lines.Name = line_name
        powers = dssCircuit.ActiveCktElement.Powers
        n = len(powers)
        P_vals = [powers[i] for i in range(0, n, 2)]
        Q_vals = [powers[i+1] for i in range(0, n, 2)]
        half = len(P_vals) // 2
        P1, P2 = sum(P_vals[:half]), sum(P_vals[half:])
        Q1, Q2 = sum(Q_vals[:half]), sum(Q_vals[half:])
        pf1 = P1 / math.sqrt(P1**2 + Q1**2) if (P1 or Q1) else 1.0
        pf2 = P2 / math.sqrt(P2**2 + Q2**2) if (P2 or Q2) else 1.0
        P_loss = P1 + P2
        Q_loss = Q1 + Q2
        return P1, Q1, P2, Q2, P_loss, Q_loss, pf1, pf2
    except Exception:
        return 0, 0, 0, 0, 0, 0, 1.0, 1.0

print("\ud83d\udd2c AN\u00c1LISIS COMPLETO: COMPARACI\u00d3N ANTES vs DESPU\u00c9S DE CARGAS NO LINEALES")
print("="*100)

# ==================================================================================
# PASO 1: AN\u00c1LISIS DEL SISTEMA ORIGINAL (SIN CARGAS NO LINEALES)
# ==================================================================================
print("\n\ud83d\udccb PASO 1: Analizando sistema ORIGINAL (SIN cargas no lineales)...")
dssText.Command = 'Clear'
dssText.Command = 'Compile "C:/Users/dpere/Downloads/13Bus/IEEE13Nodeckt.dss"'
dssText.Command = 'Solve'

# 1.1 Voltajes originales
buses_original = list(dssCircuit.AllBusNames)
voltajes_original = list(dssCircuit.AllBusVmagPu)

# 1.2 P\u00e9rdidas totales originales
losses_original = dssCircuit.Losses
P_loss_total_original = losses_original[0] / 1000
Q_loss_total_original = losses_original[1] / 1000
total_power = dssCircuit.TotalPower
pf_total_original = abs(total_power[0]) / math.sqrt(total_power[0]**2 + total_power[1]**2) if (total_power[0] or total_power[1]) else 0

# 1.3 THD original por bus
print("   Calculando THD original por bus...")
dssText.Command = 'Set mode=harmonics'
thd_original = {}
for bus in buses_original:
    thd_original[bus] = calcular_thd_bus(bus)

# 1.4 Flujos y p\u00e9rdidas por l\u00ednea originales
dssText.Command = 'Set mode=daily'
dssText.Command = 'Solve'
print("   Calculando flujos originales por l\u00ednea...")
lineas_original = list(dssCircuit.Lines.AllNames)
flujos_original = {}
for linea in lineas_original:
    P1, Q1, P2, Q2, P_loss, Q_loss, PF1, PF2 = obtener_flujos_linea(linea)
    flujos_original[linea] = {
        'P1': P1, 'Q1': Q1, 'P2': P2, 'Q2': Q2,
        'P_loss': P_loss, 'Q_loss': Q_loss,
        'PF1': PF1, 'PF2': PF2
    }

# 1.5 Desviaciones de voltaje originales
desviaciones_original = {}
for i, bus in enumerate(buses_original):
    desviacion_pct = (voltajes_original[i] - 1.0) * 100
    desviaciones_original[bus] = desviacion_pct

print("\u2705 Sistema original analizado completamente")

# ==================================================================================
# PASO 2: AN\u00c1LISIS DEL SISTEMA CON CARGAS NO LINEALES
# ==================================================================================
print("\n\ud83d\udccb PASO 2: Analizando sistema CON cargas no lineales EXTREMAS...")
dssText.Command = 'Clear'
dssText.Command = 'Compile "C:/Users/dpere/Downloads/13Bus/IEEE13Nodeckt.dss"'

# Agregar cargas no lineales EXTREMAS para hacer el impacto muy visible
print("\ud83d\udd25 Agregando cargas no lineales EXTREMAS...")

# Espectros muy agresivos
dssText.Command = 'New Spectrum.NonLinearLoad1 NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 70 60 50 45 35 30 25]'
dssText.Command = 'angle=[0 180 0 180 0 180 0 180]'

dssText.Command = 'New Spectrum.NonLinearLoad2 NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 75 65 55 50 40 35 28]'
dssText.Command = 'angle=[0 -120 120 -60 60 -120 120 -60]'

dssText.Command = 'New Spectrum.HarmonicSource NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 65 50 40 35 25 20 15]'
dssText.Command = 'angle=[0 90 -90 45 -45 90 -90 45]'

# Cargas MASIVAS (mucho m\u00e1s grandes que antes)
dssText.Command = 'New Load.NonLinear_634_1 Bus1=634 Phases=3 kW=1200 kvar=900 Model=1 Spectrum=NonLinearLoad1'
dssText.Command = 'New Load.NonLinear_634_2 Bus1=634.1 Phases=1 kW=600 kvar=450 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_634_3 Bus1=634.2 Phases=1 kW=500 kvar=375 Model=1 Spectrum=NonLinearLoad1'

dssText.Command = 'New Load.NonLinear_671 Bus1=671 Phases=3 kW=800 kvar=600 Model=1 Spectrum=HarmonicSource'
dssText.Command = 'New Load.NonLinear_645 Bus1=645.2 Phases=1 kW=400 kvar=300 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_632 Bus1=632.1 Phases=1 kW=350 kvar=250 Model=1 Spectrum=NonLinearLoad1'
dssText.Command = 'New Load.NonLinear_675 Bus1=675.1 Phases=1 kW=300 kvar=200 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_680 Bus1=680.1 Phases=1 kW=250 kvar=180 Model=1 Spectrum=HarmonicSource'
dssText.Command = 'New Load.NonLinear_652 Bus1=652.1 Phases=1 kW=200 kvar=150 Model=1 Spectrum=NonLinearLoad1'

print("\ud83d\udd25 CARGAS NO LINEALES EXTREMAS AGREGADAS:")
print("   - Bus 634: 2300 kW no lineales")
print("   - Bus 671: 800 kW no lineales")
print("   - Bus 645: 400 kW no lineales")
print("   - Otros buses: 1100 kW no lineales")
print("   - \ud83d\udca5 TOTAL: ~4600 kW no lineales (\u2248130% del sistema original)")
print("   - \ud83d\udea8 Contenido arm\u00f3nico EXTREMO (hasta 75%)")

dssText.Command = 'Solve'

# 2.1 Voltajes con cargas no lineales
buses_con_nl = list(dssCircuit.AllBusNames)
voltajes_con_nl = list(dssCircuit.AllBusVmagPu)

# 2.2 P\u00e9rdidas totales con cargas no lineales
losses_con_nl = dssCircuit.Losses
P_loss_total_con_nl = losses_con_nl[0] / 1000
Q_loss_total_con_nl = losses_con_nl[1] / 1000
total_power_nl = dssCircuit.TotalPower
pf_total_con_nl = abs(total_power_nl[0]) / math.sqrt(total_power_nl[0]**2 + total_power_nl[1]**2) if (total_power_nl[0] or total_power_nl[1]) else 0

# 2.3 THD con cargas no lineales
print("   Calculando THD con cargas no lineales...")
dssText.Command = 'Set mode=harmonics'
thd_con_nl = {}
for bus in buses_con_nl:
    thd_con_nl[bus] = calcular_thd_bus(bus)

# 2.4 Flujos y p\u00e9rdidas por l\u00ednea con cargas no lineales
dssText.Command = 'Set mode=daily'
dssText.Command = 'Solve'
print("   Calculando flujos con cargas no lineales...")
flujos_con_nl = {}
for linea in dssCircuit.Lines.AllNames:
    P1, Q1, P2, Q2, P_loss, Q_loss, PF1, PF2 = obtener_flujos_linea(linea)
    flujos_con_nl[linea] = {
        'P1': P1, 'Q1': Q1, 'P2': P2, 'Q2': Q2,
        'P_loss': P_loss, 'Q_loss': Q_loss,
        'PF1': PF1, 'PF2': PF2
    }

# 2.5 Desviaciones de voltaje con cargas no lineales
desviaciones_con_nl = {}
for i, bus in enumerate(buses_con_nl):
    desviacion_pct = (voltajes_con_nl[i] - 1.0) * 100
    desviaciones_con_nl[bus] = desviacion_pct

print("\u2705 Sistema con cargas no lineales analizado completamente")

# ==================================================================================
# PASO 3: AN\u00c1LISIS DEL SISTEMA CON FILTROS PASIVOS
# ==================================================================================
print("\n\ud83d\udee0\ufe0f PASO 3: Aplicando filtros pasivos para mejorar PF y THD...")
dssText.Command = 'Clear'
dssText.Command = 'Compile "C:/Users/dpere/Downloads/13Bus/IEEE13Nodeckt.dss"'
# Espectros agresivos (mismos del paso 2)
dssText.Command = 'New Spectrum.NonLinearLoad1 NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 70 60 50 45 35 30 25]'
dssText.Command = 'angle=[0 180 0 180 0 180 0 180]'

dssText.Command = 'New Spectrum.NonLinearLoad2 NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 75 65 55 50 40 35 28]'
dssText.Command = 'angle=[0 -120 120 -60 60 -120 120 -60]'

dssText.Command = 'New Spectrum.HarmonicSource NumHarm=8'
dssText.Command = 'harmonic=[1 3 5 7 9 11 13 15]'
dssText.Command = '%mag=[100 65 50 40 35 25 20 15]'
dssText.Command = 'angle=[0 90 -90 45 -45 90 -90 45]'

# Cargas no lineales nuevamente
dssText.Command = 'New Load.NonLinear_634_1 Bus1=634 Phases=3 kW=1200 kvar=900 Model=1 Spectrum=NonLinearLoad1'
dssText.Command = 'New Load.NonLinear_634_2 Bus1=634.1 Phases=1 kW=600 kvar=450 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_634_3 Bus1=634.2 Phases=1 kW=500 kvar=375 Model=1 Spectrum=NonLinearLoad1'

dssText.Command = 'New Load.NonLinear_671 Bus1=671 Phases=3 kW=800 kvar=600 Model=1 Spectrum=HarmonicSource'
dssText.Command = 'New Load.NonLinear_645 Bus1=645.2 Phases=1 kW=400 kvar=300 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_632 Bus1=632.1 Phases=1 kW=350 kvar=250 Model=1 Spectrum=NonLinearLoad1'
dssText.Command = 'New Load.NonLinear_675 Bus1=675.1 Phases=1 kW=300 kvar=200 Model=1 Spectrum=NonLinearLoad2'
dssText.Command = 'New Load.NonLinear_680 Bus1=680.1 Phases=1 kW=250 kvar=180 Model=1 Spectrum=HarmonicSource'
dssText.Command = 'New Load.NonLinear_652 Bus1=652.1 Phases=1 kW=200 kvar=150 Model=1 Spectrum=NonLinearLoad1'

# Filtros pasivos
dssText.Command = 'New Capacitor.Filter634 Bus1=634 Phases=3 kV=4.16 kVAr=600'
dssText.Command = 'New Capacitor.Filter671 Bus1=671 Phases=3 kV=4.16 kVAr=300'
dssText.Command = 'New Capacitor.Filter645 Bus1=645.2 Phases=1 kV=2.4 kVAr=150'
dssText.Command = 'New Reactor.HarmFilt634 Bus1=634 Phases=3 kV=4.16 kVAr=300'

dssText.Command = 'Solve'

# 3.1 Voltajes con filtros
buses_con_filt = list(dssCircuit.AllBusNames)
voltajes_con_filt = list(dssCircuit.AllBusVmagPu)

# 3.2 P\u00e9rdidas totales con filtros
losses_con_filt = dssCircuit.Losses
P_loss_total_con_filt = losses_con_filt[0] / 1000
Q_loss_total_con_filt = losses_con_filt[1] / 1000
total_power_filt = dssCircuit.TotalPower
pf_total_con_filt = abs(total_power_filt[0]) / math.sqrt(total_power_filt[0]**2 + total_power_filt[1]**2) if (total_power_filt[0] or total_power_filt[1]) else 0

# 3.3 THD con filtros
print("   Calculando THD con filtros...")
dssText.Command = 'Set mode=harmonics'
thd_con_filt = {}
for bus in buses_con_filt:
    thd_con_filt[bus] = calcular_thd_bus(bus)

# 3.4 Flujos y p\u00e9rdidas por l\u00ednea con filtros
dssText.Command = 'Set mode=daily'
dssText.Command = 'Solve'
print("   Calculando flujos con filtros...")
flujos_con_filt = {}
for linea in dssCircuit.Lines.AllNames:
    P1, Q1, P2, Q2, P_loss, Q_loss, PF1, PF2 = obtener_flujos_linea(linea)
    flujos_con_filt[linea] = {
        'P1': P1, 'Q1': Q1, 'P2': P2, 'Q2': Q2,
        'P_loss': P_loss, 'Q_loss': Q_loss,
        'PF1': PF1, 'PF2': PF2
    }

# 3.5 Desviaciones de voltaje con filtros
desviaciones_con_filt = {}
for i, bus in enumerate(buses_con_filt):
    desviacion_pct = (voltajes_con_filt[i] - 1.0) * 100
    desviaciones_con_filt[bus] = desviacion_pct

print("\u2705 Sistema con filtros analizado completamente")

# ==================================================================================
# PASO 4: COMPARACIONES DETALLADAS

# 3.1 COMPARACI\u00d3N DE VOLTAJES
print(f"\n\ud83d\udcca COMPARACI\u00d3N DETALLADA DE VOLTAJES:")
print("="*110)
print("   BUS        ORIGINAL   CON_NL   CON_FIL   CAMBIO_NL   CAMBIO_FIL   ESTADO_CAMBIO")
print("              (pu)       (pu)     (pu)      (pu)        (%)")
print("-"*110)

cambios_voltaje = []
for i, bus in enumerate(buses_original):
    if bus in buses_con_nl:
        j = buses_con_nl.index(bus)
        v_orig = voltajes_original[i]
        v_nl = voltajes_con_nl[j]
        cambio_pu = v_nl - v_orig
        cambio_pct = (cambio_pu / v_orig) * 100
        
        # Clasificar el cambio
        if abs(cambio_pct) > 5.0:
            estado = "\ud83d\udd34 EXTREMO"
        elif abs(cambio_pct) > 2.0:
            estado = "\ud83d\udd34 CR\u00cdTICO"
        elif abs(cambio_pct) > 1.0:
            estado = "\ud83d\udfe0 NOTABLE"
        elif abs(cambio_pct) > 0.5:
            estado = "\ud83d\udfe1 VISIBLE"
        elif abs(cambio_pct) > 0.1:
            estado = "\ud83d\udd35 LEVE"
        else:
            estado = "\u2705 M\u00cdNIMO"

        v_fil = voltajes_con_filt[buses_con_filt.index(bus)] if bus in buses_con_filt else v_nl
        print(f"  {bus:10s}   {v_orig:8.4f}   {v_nl:8.4f}   {v_fil:8.4f}   {cambio_pu:+8.4f}   {cambio_pct:+7.2f}%   {estado}")

        cambios_voltaje.append((bus, cambio_pct, estado))

# 3.2 COMPARACI\u00d3N DE THD
print(f"\n\ud83d\udcc8 COMPARACI\u00d3N DETALLADA DE THD:")
print("="*80)
print("   BUS        ORIGINAL   CON_NL   CON_FIL   CAMBIO_NL   ESTADO_THD")
print("              (%)        (%)      (%)        (%)")
print("-"*80)

cambios_thd = []
for bus in buses_original:
    if bus in thd_con_nl:
        thd_orig = thd_original.get(bus, 0)
        thd_nl = thd_con_nl[bus]
        thd_fil = thd_con_filt.get(bus, thd_nl)
        cambio_thd = thd_nl - thd_orig
        
        # Clasificar THD
        if thd_nl > 15:
            estado = "\ud83d\udd34 CR\u00cdTICO"
        elif thd_nl > 8:
            estado = "\ud83d\dfe0 ALTO"
        elif thd_nl > 5:
            estado = "\ud83d\dfe1 MODERADO"
        else:
            estado = "\u2705 BUENO"
        
        print(f"  {bus:10s}   {thd_orig:8.2f}   {thd_nl:8.2f}   {thd_fil:8.2f}   {cambio_thd:+8.2f}   {estado}")
        
        if cambio_thd > 1.0:
            cambios_thd.append((bus, cambio_thd, thd_nl))

# 3.3 COMPARACI\u00d3N DE FLUJOS DE POTENCIA POR L\u00cdNEA
print(f"\n\u26a1 COMPARACI\u00d3N DETALLADA DE FLUJOS DE POTENCIA POR L\u00cdNEA:")
print("="*130)
print("   L\u00cdNEA      POTENCIA ACTIVA (kW)                      POTENCIA REACTIVA (kVAr)            ESTADO")
print("              ORIG_P1    NL_P1    FIL_P1     ORIG_Q1    NL_Q1    FIL_Q1")
print("-"*130)

cambios_flujos = []
for linea in lineas_original:
    if linea in flujos_con_nl:
        # Potencia activa
        P1_orig = flujos_original[linea]['P1']
        P1_nl = flujos_con_nl[linea]['P1']
        cambio_P1 = P1_nl - P1_orig
        cambio_P1_pct = (cambio_P1 / P1_orig * 100) if P1_orig != 0 else 0
        
        # Potencia reactiva
        Q1_orig = flujos_original[linea]['Q1']
        Q1_nl = flujos_con_nl[linea]['Q1']
        cambio_Q1 = Q1_nl - Q1_orig
        cambio_Q1_pct = (cambio_Q1 / Q1_orig * 100) if Q1_orig != 0 else 0
        
        # Estado general
        if abs(cambio_P1_pct) > 10 or abs(cambio_Q1_pct) > 10:
            estado = "\ud83d\udd34 CR\u00cdTICO"
        elif abs(cambio_P1_pct) > 5 or abs(cambio_Q1_pct) > 5:
            estado = "\ud83d\dfe0 NOTABLE"
        elif abs(cambio_P1_pct) > 2 or abs(cambio_Q1_pct) > 2:
            estado = "\ud83d\dfe1 VISIBLE"
        else:
            estado = "\u2705 LEVE"
        
        P1_fil = flujos_con_filt.get(linea, {}).get('P1', P1_nl)
        Q1_fil = flujos_con_filt.get(linea, {}).get('Q1', Q1_nl)
        print(f"  {linea:10s} {P1_orig:8.1f} {P1_nl:8.1f} {P1_fil:8.1f}     {Q1_orig:8.1f} {Q1_nl:8.1f} {Q1_fil:8.1f}   {estado}")
        
        if abs(cambio_P1_pct) > 2 or abs(cambio_Q1_pct) > 2:
            cambios_flujos.append((linea, cambio_P1_pct, cambio_Q1_pct))

# 3.4 COMPARACI\u00d3N DE P\u00c9RDIDAS POR L\u00cdNEA
print(f"\n\ud83d\udd25 COMPARACI\u00d3N DETALLADA DE P\u00c9RDIDAS POR L\u00cdNEA:")
print("="*100)
print("   L\u00cdNEA      P\u00c9RDIDAS ACTIVAS (kW)                   P\u00c9RDIDAS REACTIVAS (kVAr)        ESTADO")
print("              ORIG      NL        FILT       ORIG      NL        FILT")
print("-"*100)

cambios_perdidas = []
for linea in lineas_original:
    if linea in flujos_con_nl:
        # P\u00e9rdidas activas
        P_loss_orig = flujos_original[linea]['P_loss']
        P_loss_nl = flujos_con_nl[linea]['P_loss']
        P_loss_fil = flujos_con_filt.get(linea, {}).get('P_loss', P_loss_nl)
        cambio_P_loss = P_loss_nl - P_loss_orig
        cambio_P_loss_pct = (cambio_P_loss / P_loss_orig * 100) if P_loss_orig != 0 else 0
        
        # P\u00e9rdidas reactivas
        Q_loss_orig = flujos_original[linea]['Q_loss']
        Q_loss_nl = flujos_con_nl[linea]['Q_loss']
        Q_loss_fil = flujos_con_filt.get(linea, {}).get('Q_loss', Q_loss_nl)
        cambio_Q_loss = Q_loss_nl - Q_loss_orig
        
        # Estado
        if abs(cambio_P_loss_pct) > 20:
            estado = "\ud83d\udd34 CR\u00cdTICO"
        elif abs(cambio_P_loss_pct) > 10:
            estado = "\ud83d\dfe0 ALTO"
        elif abs(cambio_P_loss_pct) > 5:
            estado = "\ud83d\dfe1 MODERADO"
        else:
            estado = "\u2705 NORMAL"
        
        print(f"  {linea:10s} {P_loss_orig:8.2f} {P_loss_nl:8.2f} {P_loss_fil:8.2f}  {Q_loss_orig:8.2f} {Q_loss_nl:8.2f} {Q_loss_fil:8.2f}  {estado}")
        
        if abs(cambio_P_loss_pct) > 5:
            cambios_perdidas.append((linea, cambio_P_loss_pct, cambio_P_loss))

# 3.5 COMPARACI\u00d3N DE DESVIACIONES DE VOLTAJE
print(f"\n\ud83d\udccc COMPARACI\u00d3N DETALLADA DE DESVIACIONES DE VOLTAJE:")
print("="*85)
print("   BUS        DESVIACI\u00d3N ORIGINAL   DESVIACI\u00d3N CON_NL   DESVIACI\u00d3N FILT   CAMBIO     ESTADO")
print("              (%)                   (%)                 (%)                (%)")
print("-"*85)

desviaciones_con_nl = desviaciones_con_nl  # already defined
cambios_desviacion = []
for bus in buses_original:
    if bus in desviaciones_con_nl:
        desv_orig = desviaciones_original[bus]
        desv_nl = desviaciones_con_nl[bus]
        desv_fil = desviaciones_con_filt.get(bus, desv_nl)
        cambio_desv = desv_nl - desv_orig
        
        # Estado de la desviaci\u00f3n
        if abs(desv_nl) > 5:
            estado = "\ud83d\udd34 FUERA L\u00cdMITE"
        elif abs(desv_nl) > 3:
            estado = "\ud83d\dfe0 CERCA L\u00cdMITE"
        elif abs(desv_nl) > 1:
            estado = "\ud83d\dfe1 MODERADO"
        else:
            estado = "\u2705 NORMAL"
        
        print(f"  {bus:10s}      {desv_orig:+8.3f}            {desv_nl:+8.3f}        {desv_fil:+8.3f}   {cambio_desv:+8.3f}   {estado}")
        
        if abs(cambio_desv) > 0.5:
            cambios_desviacion.append((bus, cambio_desv, desv_nl))

# ==================================================================================
# PASO 5: RESUMEN EJECUTIVO
# ==================================================================================
print(f"\n\ud83d\udccb RESUMEN EJECUTIVO DE IMPACTOS:")
print("="*80)

# P\u00e9rdidas totales
print(f"  \ud83d\udca5 P\u00c9RDIDAS TOTALES DEL SISTEMA:")
print(f"     Original:  {P_loss_total_original:8.2f} kW, {Q_loss_total_original:8.2f} kVAr")
print(f"     Con NL:    {P_loss_total_con_nl:8.2f} kW, {Q_loss_total_con_nl:8.2f} kVAr")
print(f"     Con Filt:  {P_loss_total_con_filt:8.2f} kW, {Q_loss_total_con_filt:8.2f} kVAr")
cambio_P_total_f = P_loss_total_con_filt - P_loss_total_original
cambio_Q_total_f = Q_loss_total_con_filt - Q_loss_total_original
cambio_P_total_f_pct = (cambio_P_total_f / P_loss_total_original) * 100
cambio_Q_total_f_pct = (cambio_Q_total_f / Q_loss_total_original) * 100
cambio_P_total = P_loss_total_con_nl - P_loss_total_original
cambio_Q_total = Q_loss_total_con_nl - Q_loss_total_original
cambio_P_total_pct = (cambio_P_total / P_loss_total_original) * 100
cambio_Q_total_pct = (cambio_Q_total / Q_loss_total_original) * 100
print(f"     Cambio:    {cambio_P_total:+8.2f} kW ({cambio_P_total_pct:+5.1f}%)")
print(f"                {cambio_Q_total:+8.2f} kVAr ({cambio_Q_total_pct:+5.1f}%)")
print(f"     Con Filt:  {cambio_P_total_f:+8.2f} kW ({cambio_P_total_f_pct:+5.1f}%)")
print(f"                {cambio_Q_total_f:+8.2f} kVAr ({cambio_Q_total_f_pct:+5.1f}%)")

print(f"\n  \u26a1 FACTOR DE POTENCIA GLOBAL:")
print(f"     Original: {pf_total_original:.3f}")
print(f"     Con NL:   {pf_total_con_nl:.3f}")
print(f"     Con Filt: {pf_total_con_filt:.3f}")

# Estad\u00edsticas de cambios
print(f"\n  \ud83d\udcca ESTAD\u00cdSTICAS DE IMPACTOS:")

# Voltajes
extremos = [x for x in cambios_voltaje if "EXTREMO" in x[2]]
criticos_v = [x for x in cambios_voltaje if "CR\u00cdTICO" in x[2]]
notables_v = [x for x in cambios_voltaje if "NOTABLE" in x[2]]
print(f"     \ud83d\udd34 Voltajes con cambios extremos (>5%): {len(extremos)} buses")
print(f"     \ud83d\udd34 Voltajes con cambios cr\u00edticos (2-5%): {len(criticos_v)} buses")
print(f"     \ud83d\dfe0 Voltajes con cambios notables (1-2%): {len(notables_v)} buses")

# THD
criticos_thd = [x for x in cambios_thd if x[2] > 15]
altos_thd = [x for x in cambios_thd if 8 < x[2] <= 15]
print(f"     \ud83d\udd34 Buses con THD cr\u00edtico (>15%): {len(criticos_thd)} buses")
print(f"     \ud83d\dfe0 Buses con THD alto (8-15%): {len(altos_thd)} buses")

# Flujos
criticos_f = [x for x in cambios_flujos if abs(x[1]) > 10 or abs(x[2]) > 10]
print(f"     \ud83d\udd34 L\u00edneas con cambios cr\u00edticos en flujos (>10%): {len(criticos_f)} l\u00edneas")

# P\u00e9rdidas
criticos_p = [x for x in cambios_perdidas if abs(x[1]) > 20]
altos_p = [x for x in cambios_perdidas if 10 < abs(x[1]) <= 20]
print(f"     \ud83d\udd34 L\u00edneas con incremento cr\u00edtico de p\u00e9rdidas (>20%): {len(criticos_p)} l\u00edneas")
print(f"     \ud83d\dfe0 L\u00edneas con incremento alto de p\u00e9rdidas (10-20%): {len(altos_p)} l\u00edneas")

# Top impactos
if cambios_voltaje:
    bus_max_v = max(cambios_voltaje, key=lambda x: abs(x[1]))
    print(f"\n  \ud83c\udfaf MAYOR IMPACTO EN VOLTAJE: {bus_max_v[0]} ({bus_max_v[1]:+.3f}%)")

if cambios_thd:
    bus_max_thd = max(cambios_thd, key=lambda x: x[2])
    print(f"  \ud83c\udfaf MAYOR THD: {bus_max_thd[0]} ({bus_max_thd[2]:.2f}%)")

if cambios_perdidas:
    linea_max_p = max(cambios_perdidas, key=lambda x: abs(x[1]))
    print(f"  \ud83c\udfaf MAYOR INCREMENTO DE P\u00c9RDIDAS: {linea_max_p[0]} ({linea_max_p[1]:+.1f}%)")

print(f"\n\ud83d\udca1 CONCLUSIONES:")
if extremos or criticos_v or criticos_thd:
    print(f"   \ud83d\udd34 IMPACTO SEVERO: Las cargas no lineales est\u00e1n causando efectos cr\u00edticos")
    print(f"   \ud83d\udd34 Se requieren medidas correctivas inmediatas")
    print(f"   \ud83d\udca1 Instalar filtros arm\u00f3nicos, mejorar factor de potencia, redistributar cargas")
elif notables_v or altos_thd:
    print(f"   \ud83d\dfe0 IMPACTO MODERADO: Las cargas no lineales est\u00e1n afectando la calidad")
    print(f"   \ud83d\dfe0 Se recomienda monitoreo continuo y medidas preventivas")
else:
    print(f"   \u2705 IMPACTO CONTROLADO: El sistema mantiene operaci\u00f3n dentro de l\u00edmites")
    print(f"   \u2705 Las cargas no lineales no comprometen la estabilidad")

print(f"\n\ud83c\udfaf RECOMENDACIONES T\u00c9CNICAS:")
print(f"   \u2022 Instalar filtros pasivos en buses con THD > 15%")
print(f"   \u2022 Considerar compensaci\u00f3n reactiva en nodos cr\u00edticos")
print(f"   \u2022 Implementar monitoreo de calidad de energ\u00eda en tiempo real")
print(f"   \u2022 Evaluar refuerzo de circuitos con mayores p\u00e9rdidas")
print(f"   \u2022 Redistribuir cargas no lineales para mejor balance del sistema")
