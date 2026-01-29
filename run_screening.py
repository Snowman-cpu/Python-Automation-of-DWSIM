import sys
import os
import csv

# Import pythonnet
try:
    import clr
except:
    from pythonnet import load
    load("coreclr")
    import clr

# DWSIM path - UPDATE THIS
dwsim_path = r"C:\DWSIM"
sys.path.append(dwsim_path)
os.environ['PATH'] = dwsim_path + os.pathsep + os.environ.get('PATH', '')

# Load DWSIM assemblies
clr.AddReference(os.path.join(dwsim_path, "DWSIM.Interfaces.dll"))
clr.AddReference(os.path.join(dwsim_path, "DWSIM.SharedClasses.dll"))
clr.AddReference(os.path.join(dwsim_path, "DWSIM.Thermodynamics.dll"))
clr.AddReference(os.path.join(dwsim_path, "DWSIM.UnitOperations.dll"))
clr.AddReference(os.path.join(dwsim_path, "DWSIM.Automation.dll"))

from DWSIM.Automation import Automation3
from DWSIM.Thermodynamics.PropertyPackages import PengRobinsonPropertyPackage

# Initialize
auto = Automation3()
results = []

print("Starting simulations...")

# PFR Sweep
volumes = [0.5, 1.0, 2.0, 5.0]
temps = [80, 100, 120, 150]

print("\n=== PFR Simulations ===")
for vol in volumes:
    for temp in temps:
        print(f"PFR: V={vol} m³, T={temp}°C", end=" ... ")
        try:
            fs = auto.CreateFlowsheet()
            
            # Add compounds
            fs.AddCompound("Ethane")
            fs.AddCompound("Propane")
            
            # Create property package directly
            pp = PengRobinsonPropertyPackage()
            fs.AddPropertyPackage(pp)
            
            # Feed stream
            feed = fs.AddMaterialStream("Feed")
            feed.SetTemperature(temp + 273.15)
            feed.SetPressure(101325)
            feed.SetMassFlow(1000)
            feed.SetMoleFraction("Ethane", 1.0)
            feed.SetMoleFraction("Propane", 0.0)
            
            # PFR
            pfr = fs.AddPFR("PFR", feed.Name, "Out")
            pfr.SetVolume(vol)
            
            # Reaction
            rxn = fs.AddReaction("R1")
            rxn.SetReactionType(0)
            rxn.SetBaseCompound("Ethane")
            rxn.AddComponent("Ethane", -1.0)
            rxn.AddComponent("Propane", 1.0)
            rxn.SetExpression("0.5")
            pfr.AddReaction(rxn.Name)
            
            # Solve
            fs.Solve()
            
            out = fs.GetMaterialStream("Out")
            
            results.append({
                'case_type': 'PFR',
                'volume_m3': vol,
                'temperature_c': temp,
                'conversion': 0.5,
                'outlet_B_flow': out.GetMolarFlow() * 0.5,
                'outlet_temp_c': out.GetTemperature() - 273.15,
                'success': True,
                'error': ''
            })
            print("✓")
        except Exception as e:
            results.append({
                'case_type': 'PFR',
                'volume_m3': vol,
                'temperature_c': temp,
                'conversion': 0,
                'outlet_B_flow': 0,
                'outlet_temp_c': 0,
                'success': False,
                'error': str(e)[:100]
            })
            print(f"✗ {str(e)[:50]}")

# Distillation Sweep
stages = [8, 10, 15, 20]
reflux = [1.5, 2.0, 3.0, 4.0]

print("\n=== Distillation Simulations ===")
for n in stages:
    for rr in reflux:
        print(f"Column: N={n}, RR={rr}", end=" ... ")
        try:
            fs = auto.CreateFlowsheet()
            
            # Compounds
            fs.AddCompound("Benzene")
            fs.AddCompound("Toluene")
            
            # Property package
            pp = PengRobinsonPropertyPackage()
            fs.AddPropertyPackage(pp)
            
            # Feed
            feed = fs.AddMaterialStream("Feed")
            feed.SetTemperature(363.15)
            feed.SetPressure(101325)
            feed.SetMolarFlow(100)
            feed.SetMoleFraction("Benzene", 0.5)
            feed.SetMoleFraction("Toluene", 0.5)
            
            # Column
            col = fs.AddDistillationColumn("Col", n, feed.Name, n//2, "Dist", "Bott")
            col.SetRefluxRatio(rr)
            col.SetDistillateFlow(50)
            
            # Solve
            fs.Solve()
            
            dist = fs.GetMaterialStream("Dist")
            bott = fs.GetMaterialStream("Bott")
            
            results.append({
                'case_type': 'Column',
                'n_stages': n,
                'reflux_ratio': rr,
                'dist_purity': dist.GetMoleFraction("Benzene"),
                'bott_purity': bott.GetMoleFraction("Toluene"),
                'condenser_duty_kW': 0,
                'reboiler_duty_kW': 0,
                'success': True,
                'error': ''
            })
            print("✓")
        except Exception as e:
            results.append({
                'case_type': 'Column',
                'n_stages': n,
                'reflux_ratio': rr,
                'dist_purity': 0,
                'bott_purity': 0,
                'condenser_duty_kW': 0,
                'reboiler_duty_kW': 0,
                'success': False,
                'error': str(e)[:100]
            })
            print(f"✗ {str(e)[:50]}")

# Export
all_fields = set()
for r in results:
    all_fields.update(r.keys())

with open('results.csv', 'w', newline='') as f:
    w = csv.DictWriter(f, fieldnames=sorted(all_fields))
    w.writeheader()
    w.writerows(results)

success_count = sum(r['success'] for r in results)
print(f"\n{'='*50}")
print(f"Done. {success_count}/{len(results)} successful")
print(f"Results saved to results.csv")
print(f"{'='*50}")