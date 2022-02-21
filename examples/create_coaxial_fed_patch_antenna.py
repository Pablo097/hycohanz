import math, cmath
import hycohanz as hfss
from hycohanz.expression import Expression as Ex

## Global variables
mu0 = 4*math.pi*1e-7
eps0 = 8.85418*1e-12
c0 = 1/math.sqrt(eps0*mu0)

## Project variables
freq_res = 14*1e9		# Hz
lambda_res = c0/freq_res

# Variables can be declared as floats in SI units, as in the following lines
## SMA coaxial connector dimensions
OuterCoaxRadius = 2.05*1e-3				# m
materialSubstrateCoax = '"Teflon (tm)"'
coaxPermittivity = 2.1					# Teflon permittivity
Z0 = 47.5                   			# ohms
# Calculate coaxial inner radius from outer radius, Z0 and permittivity
# The formula can be seen in https://www.everythingrf.com/rf-calculators/coaxial-cable-calculator
innerCoaxRadius = round(OuterCoaxRadius/math.pow(10, Z0*math.sqrt(coaxPermittivity)/138), 3)
LengthCoax = 5*1e-3						# m
thicknessOuterConductor = 180*1e-6		# m

# Variables can also be declared as strings or hycohanz Expressions, which
# is convenient when not using SI units, for example.
## Patch dimensions
thicknessSubstrate = Ex("30mils")	# mils
SubstrateSize = Ex("20mm")			# mm
materialSubstratePatch = '"Rogers RO4350 (tm)"'
thicknessPatch = Ex("180um")		# um
L = Ex("4.78mm")                   	# mm
W = Ex("7.02mm")                   	# mm
FedDealignment = Ex("1.31mm") 		# mm


projectName = "Test Project"
designName = "Test Design"

## Connect to HFSS and create project and design
hfss.setup_interface()
# Creates a new project and renames it with the name from 'projectName' string
hfss.new_project()
hfss.rename_project(hfss.get_path() + projectName + ".aedt")
# Inserts a new DrivenModal design with the named from 'designName' string
hfss.insert_design(designName, "DrivenModal")
# Sets the active editor to "3D Modeler" (The default and only known correct value)
hfss.set_active_editor()


input('Press <Enter> to save the required variables in HFSS.')

## Save dimensions as variables

# It is very convenient to have HFSS variables in order to be able to change
# them afterwards in our design.
# For example, the following function creates an HFSS variable called
# 'FedDealignment' with the value inside the Python 'FedDealignment' variable,
# i.e. "1.31mm"
hfss.add_property("FedDealignment", FedDealignment)
hfss.add_property("SubstrateSize", SubstrateSize)
hfss.add_property("thicknessSubstrate", thicknessSubstrate)
hfss.add_property("W", W)
hfss.add_property("L", L)

# If the numeric values of these variables are not necessary anymore, we can
# assign them the HFSS expression with the already created HFSS variable. Thus,
# from now on, every time we use one of these variables for any hycohanz
# function, it will leave the required value as an HFSS expression
FedDealignment = Ex("FedDealignment")
SubstrateSize = Ex("SubstrateSize")
thicknessSubstrate = Ex("thicknessSubstrate")
W = Ex("W")
L = Ex("L")

# The same can be done with the required numeric variables
hfss.add_property("LengthCoax", LengthCoax)
hfss.add_property("thicknessPatch", thicknessPatch)
LengthCoax = Ex("LengthCoax")
thicknessPatch = Ex("thicknessPatch")

# The good thing about the hycohanz Expressions is that they can be manipulated
# in Python as if they were numeric values, i.e. you can add, subtract, multiply
# divide, etc with another hycohanz Expression or any numeric value and it will
# return a new hycohanz Expression with the correct equation.
#
# For example, if we do the following:
# >> A = LengthCoax + thicknessSubstrate/2
# 'A' will be a hycohanz Expression with the following string:
# >> A: "5*1e-3 + thicknessSubstrate/2"
# which HFSS will recognize correctly. This way, you don´t have to worry about
# whether you are operating with expressions or numeric values in Python, as
# any hycohanz Expression in the operation will convert all to another expression,
# which is a valid input to any hfss function.


input('Press <Enter> to create the GND plane.')

## Make GND Plane
# As 'SubstrateSize' is a hycohanz Expression, '-SubstrateSize/2' will become
# a hycohanz Expression with the string "-SubstrateSize/2" inside.
hfss.create_rectangle(-SubstrateSize/2, -SubstrateSize/2, 0,
						SubstrateSize, SubstrateSize,	Name="GNDPlane",
						Transparency = 0.2, Color=(200,143,14))
# Assign Perfect E boundary to the created plane
hfss.assign_perfect_e("PerfE_GND", hfss.get_face_ids("GNDPlane"))


input('Press <Enter> to create the substrate.')

## Create substrate Plane
hfss.create_box(-SubstrateSize/2, -SubstrateSize/2, 0,
				SubstrateSize, SubstrateSize, thicknessSubstrate,
				MaterialValue = materialSubstratePatch, Name="SubstratePlane",
				Transparency = 0.7)

input('Press <Enter> to create the patch.')

## Create Patch
hfss.create_box(-W/2, -L/2, thicknessSubstrate,
				W, L, thicknessPatch, MaterialValue = '"copper"', Name="Patch",
				Transparency = 0.1, Color=(200,143,14))

input('Press <Enter> to create the coaxial cable and probe.')

## Substract coaxial input from GND Plane
hfss.create_circle(0, -FedDealignment, 0, OuterCoaxRadius, Name="GNDHole")
hfss.subtract(["GNDPlane"], ["GNDHole"])

## Create Coaxial exterior conductive layer
hfss.create_cylinder(0, -FedDealignment, -LengthCoax,
					OuterCoaxRadius+thicknessOuterConductor, LengthCoax,
					MaterialValue='"pec"', Name="Coaxial")
hfss.create_cylinder(0, -FedDealignment, -LengthCoax,
					OuterCoaxRadius, LengthCoax, Name="CoaxialErase")
hfss.subtract(["Coaxial"], ["CoaxialErase"])

## Create Coaxial substrate
hfss.create_cylinder(0, -FedDealignment, -LengthCoax,
					OuterCoaxRadius, LengthCoax,
					MaterialValue=materialSubstrateCoax, Name="CoaxialInterior")
hfss.create_cylinder(0, -FedDealignment, -LengthCoax,
					innerCoaxRadius, LengthCoax, Name="CoaxialInteriorErase")
hfss.subtract(["CoaxialInterior"], ["CoaxialInteriorErase"])

## Create Coaxial core
hfss.create_cylinder(0, -FedDealignment, -LengthCoax,
					innerCoaxRadius, LengthCoax,
					MaterialValue='"pec"', Name="CoaxialCore")

## Create Coaxial probe
hfss.create_cylinder(0, -FedDealignment, 0,
					innerCoaxRadius, thicknessSubstrate,
					MaterialValue='"pec"', Name="CoaxialProbe")
# Subtract coaxial probe from Substrate Plane
hfss.subtract(["SubstratePlane"], ["CoaxialProbe"], KeepOriginals=True)

input('Press <Enter> to assign the excitation to the coaxial cable.')

## Obtain the face in which the excitation must be applied
waveport_face_id = hfss.get_face_by_position("CoaxialInterior",
			(OuterCoaxRadius+innerCoaxRadius)/2, -FedDealignment, -LengthCoax)

## Create the Lumped Port excitation with an integration line from the coaxial
## inner layer to the outer one.
hfss.assign_lumpedport_intline([waveport_face_id],
					[0, -FedDealignment-innerCoaxRadius, -LengthCoax],
					[0, -FedDealignment-OuterCoaxRadius, -LengthCoax],
					Impedance=str(Z0)+"ohm", portname="CoaxialWaveport")
# As can be notices, 'FedDealignment' and 'LengthCoax' are expressions, and HFSS
# doesn´t let the start and end coordinates of the port integration lines to be
# expressions in normal circunstances. Luckily, the hycohanz library takes care
# of this problem and lets us use common HFSS expressions and it will continue
# working as usual!


input('Press <Enter> to create the radiation region.')

hfss.create_open_region(13*1e9)

input('Press <Enter> to create the analysis setup and frequency sweep.')

## Analysis setup
freqMax = 16*1e9		# Hz
freqMin = 12*1e9		# Hz
freqStep = 0.1*1e9;		# Hz
minDeltaS = 0.01;

analysisName = hfss.insert_analysis_setup(freqMax, minDeltaS, MinimumPasses=10, MinimumConvergedPasses=3)
hfss.insert_frequency_sweep(analysisName, freqMin, freqMax, freqStep, SaveRadFieldsOnly=False)

# We could solve the design with the following function
# hfss.solve([analysisName])

input('Press <Enter> to save the project, clean the interface and end the script.')

hfss.save_project()
hfss.clean_interface()
