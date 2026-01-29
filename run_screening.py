import sys
import os
import csv
import traceback
from datetime import datetime

# ---- CONFIG ----

DWSIM_PATH = r"C:\DWSIM"
sys.path.append(DWSIM_PATH)
os.environ["PATH"] = DWSIM_PATH + os.pathsep + os.environ["PATH"]

import clr
clr.AddReferenceToFileAndPath(os.path.join(DWSIM_PATH, "ThermoCS.dll"))
clr.AddReference("DWSIM.Automation")

from DWSIM.Automation import Automation3

RESULTS_FILE = "results.csv"


def init_results_csv():
    if not os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([
                "timestamp",
                "case_type",
                "success",
                "error_message",

                # PFR sweep vars
                "pfr_volume",
                "pfr_temperature",
                "pfr_conversion",
                "pfr_B_outlet_flow",
                "pfr_heat_duty",
                "pfr_outlet_temp",

                # Column sweep vars
                "reflux_ratio",
                "num_stages",
                "distillate_purity",
                "bottoms_purity",
                "condenser_duty",
                "reboiler_duty"
            ])


def log_result(row):
    with open(RESULTS_FILE, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(row)


def run_pfr_case(auto, volume, temperature):
    """
    Builds flowsheet + PFR programmatically
    """
    fs = auto.CreateFlowsheet()

    try:
        # ---- Create Components ----
        fs.AddCompound("Methanol")   # A
        fs.AddCompound("Ethanol")    # B

        # ---- Material Stream ----
        feed = fs.AddMaterialStream("Feed")
        feed.SetTemperature(temperature)
        feed.SetPressure(101325)
        feed.SetMolarFlow(100)
        feed.SetOverallComposition({"Methanol": 1.0})

        # ---- PFR ----
        pfr = fs.AddUnitOp("Plug Flow Reactor", "PFR-1")
        pfr.SetPropertyValue("Volume", volume)
        pfr.SetPropertyValue("Isothermal", True)

        fs.ConnectStreams(feed, pfr)

        # ---- Reaction (A -> B) ----
        rxn = fs.AddReaction("R1")
        rxn.SetReactionType("Kinetic")
        rxn.SetStoichiometry({"Methanol": -1, "Ethanol": 1})
        rxn.SetRateExpression("k*C_Methanol")
        rxn.SetParameter("k", 0.1)

        fs.AddReactionToSet(rxn)

        # ---- Run Simulation ----
        fs.Solve()

        outlet = pfr.GetOutletMaterialStream()

        conversion = 1 - outlet.GetMolarFraction("Methanol")
        B_flow = outlet.GetComponentMolarFlow("Ethanol")
        heat_duty = pfr.GetPropertyValue("Heat Duty")
        outlet_temp = outlet.GetTemperature()

        return {
            "conversion": conversion,
            "B_flow": B_flow,
            "heat_duty": heat_duty,
            "outlet_temp": outlet_temp
        }

    finally:
        auto.CloseFlowsheet(fs)


def run_column_case(auto, reflux_ratio, num_stages):
    fs = auto.CreateFlowsheet()

    try:
        # ---- Components ----
        fs.AddCompound("Benzene")
        fs.AddCompound("Toluene")

        # ---- Feed ----
        feed = fs.AddMaterialStream("Feed")
        feed.SetTemperature(360)
        feed.SetPressure(101325)
        feed.SetMolarFlow(100)
        feed.SetOverallComposition({
            "Benzene": 0.5,
            "Toluene": 0.5
        })

        # ---- Distillation Column ----
        col = fs.AddUnitOp("Distillation Column", "COL-1")
        col.SetPropertyValue("Number of Stages", num_stages)
        col.SetPropertyValue("Reflux Ratio", reflux_ratio)
        col.SetPropertyValue("Feed Stage", int(num_stages/2))
        col.SetPropertyValue("Condenser Type", "Total")

        fs.ConnectStreams(feed, col)

        # ---- Solve ----
        fs.Solve()

        dist = col.GetDistillateStream()
        bot = col.GetBottomsStream()

        benzene_dist = dist.GetMolarFraction("Benzene")
        toluene_bot = bot.GetMolarFraction("Toluene")

        condenser_duty = col.GetPropertyValue("Condenser Duty")
        reboiler_duty = col.GetPropertyValue("Reboiler Duty")

        return {
            "dist_purity": benzene_dist,
            "bot_purity": toluene_bot,
            "cond_duty": condenser_duty,
            "reb_duty": reboiler_duty
        }

    finally:
        auto.CloseFlowsheet(fs)


def main():
    init_results_csv()
    auto = Automation3()

    # ---- PFR PARAMETRIC SWEEP ----
    pfr_volumes = [1, 5, 10]
    pfr_temps = [350, 375, 400]

    for V in pfr_volumes:
        for T in pfr_temps:
            try:
                res = run_pfr_case(auto, V, T)
                log_result([
                    datetime.now(), "PFR", True, "",
                    V, T,
                    res["conversion"],
                    res["B_flow"],
                    res["heat_duty"],
                    res["outlet_temp"],
                    "", "", "", "", "", ""
                ])
            except Exception as e:
                log_result([
                    datetime.now(), "PFR", False, str(e),
                    V, T, "", "", "", "",
                    "", "", "", "", "", ""
                ])

    # ---- COLUMN PARAMETRIC SWEEP ----
    reflux_ratios = [1.5, 2.0, 3.0]
    stages_list = [10, 15, 20]

    for RR in reflux_ratios:
        for N in stages_list:
            try:
                res = run_column_case(auto, RR, N)
                log_result([
                    datetime.now(), "DIST", True, "",
                    "", "", "", "", "", "",
                    RR, N,
                    res["dist_purity"],
                    res["bot_purity"],
                    res["cond_duty"],
                    res["reb_duty"]
                ])
            except Exception as e:
                log_result([
                    datetime.now(), "DIST", False, str(e),
                    "", "", "", "", "", "",
                    RR, N, "", "", "", ""
                ])

    print("✔ Screening automation completed successfully.")


if __name__ == "__main__":
    main()
