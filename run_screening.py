"""
DWSIM Automation Script - Screening Task 2 (FIXED VERSION)
Simulates PFR reactor and distillation column with parametric sweeps
"""

import sys
import os
import csv
from pathlib import Path
import traceback

# Fix pythonnet import - ensure we get the correct clr module
try:
    from pythonnet import load
    load("coreclr")
    import clr
except:
    import clr

# Verify clr has AddReference
if not hasattr(clr, 'AddReference'):
    print("ERROR: Wrong 'clr' module loaded. Fixing...")
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", "clr"])
    subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "--force-reinstall", "pythonnet"])
    print("Please run the script again after pythonnet reinstallation")
    sys.exit(1)

# ============================================================================
# CONFIGURATION - UPDATE THIS PATH TO YOUR DWSIM INSTALLATION
# ============================================================================
dwsim_path = r"C:\DWSIM"  # YOUR DWSIM PATH (found it!)
# ============================================================================

# Add DWSIM path to system path
sys.path.append(dwsim_path)
os.environ['PATH'] = dwsim_path + os.pathsep + os.environ['PATH']

# Load all required DWSIM assemblies with full paths
try:
    print("Loading DWSIM assemblies...")
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.Automation.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.Interfaces.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.GlobalSettings.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.SharedClasses.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.Thermodynamics.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.UnitOperations.dll"))
    clr.AddReference(os.path.join(dwsim_path, "DWSIM.Inspector.dll"))
    clr.AddReference(os.path.join(dwsim_path, "CapeOpen.dll"))
    print("✓ All assemblies loaded successfully")
except Exception as e:
    print(f"✗ Error loading assemblies: {e}")
    print(f"   Make sure DWSIM is installed at: {dwsim_path}")
    sys.exit(1)

from DWSIM.Automation import Automation3
from DWSIM.Interfaces.Enums.GraphicObjects import ObjectType


class DWSIMAutomation:
    """Handles DWSIM automation and simulation tasks"""
    
    def __init__(self):
        self.interf = Automation3()
        self.flowsheet = None
        self.results = []
        
    def create_flowsheet(self):
        """Create a new flowsheet (no name parameter needed)"""
        try:
            # Create flowsheet without parameters - API doesn't accept string name
            self.flowsheet = self.interf.CreateFlowsheet()
            return True
        except Exception as e:
            print(f"Error creating flowsheet: {e}")
            traceback.print_exc()
            return False
    
    def setup_compounds(self, compound_list):
        """Add compounds to the flowsheet"""
        try:
            # Method 1: Try using compound database
            try:
                compound_db = self.interf.GetCompounds()
                for compound_name in compound_list:
                    if compound_name in compound_db.Keys:
                        comp_obj = compound_db[compound_name]
                        self.flowsheet.SelectedCompounds.Add(comp_obj.Name, comp_obj)
                return True
            except Exception as e1:
                print(f"  Method 1 (GetCompounds) failed: {e1}")
            
            # Method 2: Try QTFillCompoundsList
            try:
                self.interf.QTFillCompoundsList(self.flowsheet)
                for compound_name in compound_list:
                    comp_obj = self.flowsheet.AvailableCompounds[compound_name]
                    self.flowsheet.SelectedCompounds.Add(comp_obj.Name, comp_obj)
                return True
            except Exception as e2:
                print(f"  Method 2 (QTFillCompoundsList) failed: {e2}")
            
            # Method 3: Try AddComponent if it exists
            try:
                for compound_name in compound_list:
                    self.flowsheet.AddComponent(compound_name)
                return True
            except Exception as e3:
                print(f"  Method 3 (AddComponent) failed: {e3}")
            
            # If all methods fail
            raise Exception("All compound addition methods failed")
            
        except Exception as e:
            print(f"Error adding compounds: {e}")
            traceback.print_exc()
            return False
    
    def setup_property_package(self, pkg_name="Peng-Robinson (PR)"):
        """Setup thermodynamic property package"""
        try:
            # Create property package
            pp = self.interf.CreatePropertyPackage(pkg_name)
            self.flowsheet.PropertyPackages.Add(pp.UniqueID, pp)
            return True
        except Exception as e:
            print(f"Error setting property package: {e}")
            traceback.print_exc()
            return False
    
    def simulate_pfr(self, volume_m3=1.0, temperature_c=100.0, pressure_bar=1.0):
        """
        Simulate PFR reactor with reaction A -> B
        """
        result = {
            'case_type': 'PFR',
            'volume_m3': volume_m3,
            'temperature_c': temperature_c,
            'pressure_bar': pressure_bar,
            'success': False,
            'error': None
        }
        
        try:
            # Create fresh flowsheet
            if not self.create_flowsheet():
                result['error'] = "Failed to create flowsheet"
                return result
            
            # Add compounds (using simple hydrocarbons as proxy for A, B)
            if not self.setup_compounds(["Ethane", "Propane"]):
                result['error'] = "Failed to add compounds"
                return result
                
            if not self.setup_property_package("Peng-Robinson (PR)"):
                result['error'] = "Failed to setup property package"
                return result
            
            # Create material stream (feed)
            feed = self.flowsheet.AddObject(ObjectType.MaterialStream, "Feed")
            
            # Set feed properties using the stream's property methods
            feed.PropertyPackage = list(self.flowsheet.PropertyPackages.Values)[0]
            
            # Set temperature, pressure, and flow
            feed.SetPropertyValue("Temperature", temperature_c + 273.15)
            feed.SetPropertyValue("Pressure", pressure_bar * 101325)
            feed.SetPropertyValue("MassFlow", 1000.0)
            
            # Set composition (100% component A - Ethane)
            feed.SetPropertyValue("Compounds", ["Ethane", "Propane"])
            feed.SetPropertyValue("MoleFractions", [1.0, 0.0])
            
            # Flash the feed to calculate properties
            feed.Calculate()
            
            # Create PFR reactor
            pfr = self.flowsheet.AddObject(ObjectType.RCT_PFR, "PFR")
            
            # Set reactor volume
            pfr.SetPropertyValue("Volume", volume_m3)
            pfr.SetPropertyValue("ReactorOperationMode", 0)  # Isothermal
            
            # Create outlet stream
            outlet = self.flowsheet.AddObject(ObjectType.MaterialStream, "Outlet")
            outlet.PropertyPackage = list(self.flowsheet.PropertyPackages.Values)[0]
            
            # Connect streams
            self.flowsheet.ConnectObjects(feed.Name, pfr.Name, 0, 0)
            self.flowsheet.ConnectObjects(pfr.Name, outlet.Name, 0, 0)
            
            # Add simple conversion reaction
            rxn = self.flowsheet.AddReaction("Conversion")
            rxn.ReactionType = 0  # Conversion
            rxn.BaseReactant = "Ethane"
            rxn.ReactionBasis = 0  # Molar basis
            
            # Set stoichiometry
            rxn.Components.Clear()
            rxn.Components.Add("Ethane", -1.0)
            rxn.Components.Add("Propane", 1.0)
            
            # Set conversion (50%)
            rxn.Expression = "0.5"
            
            # Assign reaction to reactor
            pfr.Reactions.Add(rxn.ID)
            
            # Calculate flowsheet
            self.flowsheet.SolveFlowsheet()
            
            # Extract results
            try:
                conversion = 0.5  # From reaction
                outlet_b_flow = outlet.GetPropertyValue("MolarFlow") * outlet.GetPropertyValue("MoleFractions")[1]
                outlet_temp = outlet.GetPropertyValue("Temperature") - 273.15
                heat_duty = pfr.GetPropertyValue("DeltaQ") / 1000.0  # W to kW
                outlet_pressure = outlet.GetPropertyValue("Pressure") / 101325  # Pa to atm
                
                result.update({
                    'success': True,
                    'conversion': conversion,
                    'outlet_B_flow_mol_s': outlet_b_flow / 3600.0,  # kmol/h to mol/s
                    'outlet_temperature_c': outlet_temp,
                    'heat_duty_kW': heat_duty,
                    'outlet_pressure_bar': outlet_pressure
                })
            except Exception as e:
                result['error'] = f"Error extracting results: {e}"
                result['traceback'] = traceback.format_exc()
            
        except Exception as e:
            result['error'] = str(e)
            result['traceback'] = traceback.format_exc()
            
        return result
    
    def simulate_distillation(self, n_stages=10, feed_stage=5, reflux_ratio=2.0, 
                            distillate_rate_kmol_h=50.0):
        """
        Simulate binary distillation column
        """
        result = {
            'case_type': 'Distillation',
            'n_stages': n_stages,
            'feed_stage': feed_stage,
            'reflux_ratio': reflux_ratio,
            'distillate_rate': distillate_rate_kmol_h,
            'success': False,
            'error': None
        }
        
        try:
            # Create fresh flowsheet
            if not self.create_flowsheet():
                result['error'] = "Failed to create flowsheet"
                return result
            
            # Add binary mixture
            if not self.setup_compounds(["Benzene", "Toluene"]):
                result['error'] = "Failed to add compounds"
                return result
                
            if not self.setup_property_package("Peng-Robinson (PR)"):
                result['error'] = "Failed to setup property package"
                return result
            
            # Create feed stream
            feed = self.flowsheet.AddObject(ObjectType.MaterialStream, "Feed")
            feed.PropertyPackage = list(self.flowsheet.PropertyPackages.Values)[0]
            
            # Set feed conditions
            feed.SetPropertyValue("Temperature", 363.15)  # 90°C
            feed.SetPropertyValue("Pressure", 101325)  # 1 atm
            feed.SetPropertyValue("MolarFlow", 100.0)  # kmol/h
            feed.SetPropertyValue("Compounds", ["Benzene", "Toluene"])
            feed.SetPropertyValue("MoleFractions", [0.5, 0.5])
            
            # Flash feed
            feed.Calculate()
            
            # Create distillation column
            column = self.flowsheet.AddObject(ObjectType.DistillationColumn, "Column")
            
            # Configure column
            column.SetPropertyValue("NumberOfStages", n_stages)
            column.SetPropertyValue("CondenserType", 0)  # Total condenser
            column.SetPropertyValue("ReboilerType", 0)  # Kettle reboiler
            
            # Set specifications
            column.SetPropertyValue("RefluxRatio", reflux_ratio)
            column.SetPropertyValue("ProductMolarFlowSpec1", distillate_rate_kmol_h)
            
            # Create product streams
            distillate = self.flowsheet.AddObject(ObjectType.MaterialStream, "Distillate")
            bottoms = self.flowsheet.AddObject(ObjectType.MaterialStream, "Bottoms")
            
            distillate.PropertyPackage = list(self.flowsheet.PropertyPackages.Values)[0]
            bottoms.PropertyPackage = list(self.flowsheet.PropertyPackages.Values)[0]
            
            # Connect streams
            self.flowsheet.ConnectObjects(feed.Name, column.Name, 0, feed_stage - 1)
            self.flowsheet.ConnectObjects(column.Name, distillate.Name, 0, 0)
            self.flowsheet.ConnectObjects(column.Name, bottoms.Name, 1, 0)
            
            # Solve flowsheet
            self.flowsheet.SolveFlowsheet()
            
            # Extract results
            try:
                dist_comp = distillate.GetPropertyValue("MoleFractions")
                bot_comp = bottoms.GetPropertyValue("MoleFractions")
                
                result.update({
                    'success': True,
                    'distillate_purity_light': dist_comp[0] if len(dist_comp) > 0 else 0.0,
                    'bottoms_purity_heavy': bot_comp[1] if len(bot_comp) > 1 else 0.0,
                    'condenser_duty_kW': abs(column.GetPropertyValue("CondenserDuty") / 1000.0),
                    'reboiler_duty_kW': abs(column.GetPropertyValue("ReboilerDuty") / 1000.0),
                    'condenser_temp_c': column.GetPropertyValue("CondenserTemperature") - 273.15
                })
            except Exception as e:
                result['error'] = f"Error extracting results: {e}"
                result['traceback'] = traceback.format_exc()
            
        except Exception as e:
            result['error'] = str(e)
            result['traceback'] = traceback.format_exc()
            
        return result
    
    def run_pfr_parametric_sweep(self):
        """Run parametric sweep for PFR"""
        print("\n=== Running PFR Parametric Sweep ===")
        
        volumes = [0.5, 1.0, 2.0, 5.0]  # m3
        temperatures = [80, 100, 120, 150]  # °C
        
        for vol in volumes:
            for temp in temperatures:
                print(f"Simulating PFR: Volume={vol} m³, Temp={temp}°C")
                result = self.simulate_pfr(volume_m3=vol, temperature_c=temp)
                self.results.append(result)
                
                if result['success']:
                    print(f"  ✓ Success - Conversion: {result['conversion']:.2%}")
                else:
                    print(f"  ✗ Failed - {result.get('error', 'Unknown error')}")
    
    def run_distillation_parametric_sweep(self):
        """Run parametric sweep for distillation column"""
        print("\n=== Running Distillation Parametric Sweep ===")
        
        stages = [8, 10, 15, 20]
        reflux_ratios = [1.5, 2.0, 3.0, 4.0]
        
        for n_stage in stages:
            for rr in reflux_ratios:
                feed_stage = max(3, n_stage // 2)  # Feed at middle
                print(f"Simulating Column: Stages={n_stage}, RR={rr}")
                result = self.simulate_distillation(
                    n_stages=n_stage, 
                    feed_stage=feed_stage,
                    reflux_ratio=rr
                )
                self.results.append(result)
                
                if result['success']:
                    purity = result.get('distillate_purity_light', 0)
                    print(f"  ✓ Success - Dist Purity: {purity:.2%}")
                else:
                    print(f"  ✗ Failed - {result.get('error', 'Unknown error')}")
    
    def export_results(self, filename="results.csv"):
        """Export results to CSV file"""
        if not self.results:
            print("No results to export")
            return
        
        # Get all possible keys
        all_keys = set()
        for result in self.results:
            all_keys.update(result.keys())
        
        # Sort keys for consistent output
        fieldnames = sorted(all_keys)
        
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for result in self.results:
                writer.writerow(result)
        
        print(f"\n✓ Results exported to {filename}")
        print(f"  Total cases: {len(self.results)}")
        print(f"  Successful: {sum(1 for r in self.results if r['success'])}")
        print(f"  Failed: {sum(1 for r in self.results if not r['success'])}")


def main():
    """Main execution function"""
    print("=" * 60)
    print("DWSIM Automation - Screening Task 2")
    print("=" * 60)
    
    try:
        # Initialize automation
        dwsim = DWSIMAutomation()
        
        # Run parametric sweeps
        dwsim.run_pfr_parametric_sweep()
        dwsim.run_distillation_parametric_sweep()
        
        # Export results
        dwsim.export_results("results.csv")
        
        print("\n" + "=" * 60)
        print("Simulation Complete!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n✗ Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()