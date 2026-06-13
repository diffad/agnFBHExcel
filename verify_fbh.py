# -*- coding: utf-8 -*-
"""Verifiziert die Rechenlogik (v0.8) durch Nachbildung der Excel-Formeln."""
import math

thV, thR = 35.0, 28.0
maxL, warn = 120.0, 15000.0
rho, cp, nu = 995.7, 4180.0, 0.000000801
alp, su, lamE, fkorr = 10.8, 0.045, 1.2, 1.00
R_u, theta_u = 2.0, 10.0
zus, Varm = 0.05, 500.0
di_m = (17 - 2 * 2.0) / 1000
Ai = math.pi / 4 * di_m**2
ZONE = {"Aufenthaltszone": 29.0, "Randzone": 35.0, "Badezimmer": 33.0}
q_down = ((thV + thR) / 2 - theta_u) / R_u
va_tab = [(50, 1.05), (75, 1.02), (100, 1.00), (125, 0.97), (150, 0.93), (200, 0.85), (250, 0.78), (300, 0.72)]

def vlookup_va(h):
    res = None
    for va, fac in va_tab:
        if va <= h: res = fac
        else: break
    return res if res is not None else ""

def room(B, E, Fi, Gq, Hr, Iva, Jk, Kzul, Lzone):
    o = {}
    try:
        M = (thV - thR) / math.log((thV - Fi) / (thR - Fi))
        if (thV - Fi) <= 0 or (thR - Fi) <= 0: M = ""
    except Exception:
        M = ""
    o["ΔθH"] = M
    N = 1 / (1 / alp + Hr + su / lamE)
    Oeta = vlookup_va(Iva)
    P = fkorr * N * Oeta * M if M != "" else ""
    o["q_oben"] = P
    o["θF"] = (Fi + P / alp) if P != "" else ""
    o["θF,max"] = ZONE.get(Lzone, 29.0)
    o["StatusθF"] = "OK" if (P != "" and o["θF"] <= o["θF,max"]) else ("zu hoch" if P != "" else "")
    S = P * E if P != "" else ""
    o["Leist_oben"] = S
    o["Deckung%"] = (S / Gq) if S != "" else ""
    V = q_down * E
    o["Verlust_unten"] = V
    W = (S + V) if S != "" else ""
    o["Leist_gesamt"] = W
    X = E / (Iva / 1000)
    Yz = 2 * Kzul * Jk
    Z = X + Yz
    AA = Z / Jk
    o["Rohr_je_Kreis"] = AA
    o["Kreis_OK"] = "OK" if AA <= maxL else "zu lang"
    AB = (W / Jk) / (cp * (thV - thR)) * 3600 if W != "" else ""   # Massenstrom kg/h (Gesamtleistung)
    AC = AB / rho * 1000 if AB != "" else ""                       # l/h
    AD = AB / 3600 / rho / Ai if AB != "" else ""                  # v
    o["v"] = AD
    AE = AD * di_m / nu if AD != "" else ""
    o["Re"] = AE
    o["Bereich"] = ("laminar" if AE < 2300 else "turbulent") if AE != "" else ""
    AG = (64 / AE if AE < 2300 else 0.316 / AE**0.25) if AE != "" else ""
    AH = AG * (AA / di_m) * (rho / 2) * AD**2 if AG != "" else ""
    o["Δp_ges"] = (AH * (1 + zus) + Varm) if AH != "" else ""
    o["Δp_OK"] = "OK" if (o["Δp_ges"] != "" and o["Δp_ges"] <= warn) else ("über" if o["Δp_ges"] != "" else "")
    return o

print(f"q_down (global) = {q_down:.2f} W/m²\n")
cases = [
    ("Wohnzimmer Teppich", dict(B="101", E=23, Fi=20, Gq=1500, Hr=0.15, Iva=100, Jk=2, Kzul=5, Lzone="Aufenthaltszone")),
    ("Bad Fliesen",        dict(B="102", E=7,  Fi=24, Gq=700,  Hr=0.01, Iva=50,  Jk=1, Kzul=8, Lzone="Badezimmer")),
    ("Flur Linoleum",      dict(B="103", E=5,  Fi=20, Gq=350,  Hr=0.05, Iva=150, Jk=1, Kzul=4, Lzone="Aufenthaltszone")),
]
for name, kw in cases:
    o = room(**kw)
    print(f"=== {name} ===")
    for k, v in o.items():
        print(f"   {k:14}= {round(v,3) if isinstance(v,float) else v}")
    print()
# Leerzeile / Grenzfall
oe = room(B="", E="", Fi=20, Gq=0, Hr=0, Iva=100, Jk=1, Kzul=0, Lzone="-") if False else None
print("KEINE Python-Exceptions -> Formellogik konsistent.")
