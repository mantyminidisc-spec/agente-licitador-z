# -*- coding: utf-8 -*-
"""
===============================================================================
LICITAIA · PLANIFICADOR INTERNO (v11.2)
===============================================================================
Ejecuta el agente en los horarios definidos (hora de Chile) SIN depender de
Coolify ni del nombre del contenedor.

Por qué existe: la tarea programada de Coolify apunta a un contenedor cuyo
nombre cambia en cada redeploy; si se olvida actualizarlo, el agente no corre.
Este planificador vive dentro del contenedor y sobrevive a los redeploys.

Horario por variable de entorno (por defecto 9,11,13,15,17,19):
    HORARIO_ESCANEO=9,11,13,15,17,19

Cada ejecución se registra en /data/planificador.log
===============================================================================
"""
import datetime
import json
import os
import subprocess
import sys
import time

RAIZ = os.getenv("LICITAIA_DATA_DIR") or "/data"
AGENTE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "agente_licitador_z.py")
ESTADO = os.path.join(RAIZ, "planificador_estado.json")
REGISTRO = os.path.join(RAIZ, "planificador.log")
MINUTO_EJECUCION = int(os.getenv("PLANIFICADOR_MINUTO", "5"))


def _horas():
    crudo = os.getenv("HORARIO_ESCANEO", "9,11,13,15,17,19")
    horas = []
    for parte in crudo.replace(";", ",").split(","):
        parte = parte.strip()
        if parte.isdigit() and 0 <= int(parte) <= 23:
            horas.append(int(parte))
    return sorted(set(horas)) or [9, 11, 13, 15, 17, 19]


def _ahora_chile():
    try:
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("America/Santiago"))
    except Exception:
        return datetime.datetime.now()


def _leer_estado():
    try:
        with open(ESTADO, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_estado(datos):
    try:
        os.makedirs(RAIZ, exist_ok=True)
        with open(ESTADO, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception as e:
        _log(f"No se pudo guardar el estado: {e}")


def _log(mensaje):
    linea = f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} · {mensaje}"
    print(linea, flush=True)
    try:
        os.makedirs(RAIZ, exist_ok=True)
        with open(REGISTRO, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except Exception:
        pass


def _ejecutar():
    if not os.path.isfile(AGENTE):
        _log(f"ERROR: no se encontró el agente en {AGENTE}")
        return False
    _log("Ejecutando el agente (--auto)")
    try:
        resultado = subprocess.run(
            [sys.executable, AGENTE, "--auto"],
            cwd=os.path.dirname(AGENTE),
            timeout=3500,
        )
        _log(f"Agente terminado con código {resultado.returncode}")
        return resultado.returncode == 0
    except subprocess.TimeoutExpired:
        _log("El agente superó el tiempo máximo (3500s) y se interrumpió")
    except Exception as e:
        _log(f"ERROR al ejecutar el agente: {e}")
    return False


def main():
    horas = _horas()
    _log(f"Planificador iniciado · horario Chile: {', '.join(f'{h:02d}:00' for h in horas)} · "
         f"minuto de ejecución: {MINUTO_EJECUCION}")
    while True:
        try:
            ahora = _ahora_chile()
            if ahora.hour in horas and ahora.minute >= MINUTO_EJECUCION:
                marca = f"{ahora:%Y-%m-%d}-{ahora.hour:02d}"
                estado = _leer_estado()
                # NO se marca antes de ejecutar: el agente es quien deja la marca
                # al arrancar. Si marcáramos aquí, el agente se saltaría su propia
                # corrida ("ya se ejecutó esta hora").
                ya_corrio = estado.get("ultima_marca") == marca
                ya_intento = estado.get("ultimo_intento") == marca
                if not ya_corrio and not ya_intento:
                    _log(f"Lanzando el agente para la corrida de las {ahora:%H}:00")
                    _ejecutar()
                    estado = _leer_estado()
                    estado["ultimo_intento"] = marca
                    _guardar_estado(estado)
                    time.sleep(300)
                else:
                    time.sleep(60)
            else:
                time.sleep(30)
        except KeyboardInterrupt:
            _log("Planificador detenido")
            return 0
        except Exception as e:
            _log(f"Error inesperado: {e}")
            time.sleep(60)


if __name__ == "__main__":
    sys.exit(main())
