# CSIOAPI

import comtypes.client

APIPath = 'C:\\Users\kgond\Documents\\API_SAP2000\\Pro_Analysis_Model\\4-storey.sdb'

ETABSObjective = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")

# create SapModel object
SapModel = ETABSObjective.Sapmodel

# initialize model
SapModel.InitializeNewModel(6)

# Set a new file
ret = SapModel.File.NewBlank()

# create model from template
# ret = SapModel.File.New2DFrame(0, 4, 3.5, 4, 4, True, 'B1', 'C1')

# define the add quick materials
ret = SapModel.PropMaterial.AddQuick('Name', 2,0,14,0,0, 0)
ret = SapModel.PropMaterial.AddQuick('Name', 6,0,0,0,0, 10)

# define material property
#MATERIAL_CONCRETE = 2
#ret = SapModel.PropMaterial.SetMaterial('CONC', 2)
# ret = SapModel.PropMaterial.SetWeightAndMass('CONC', 1, 25)
# assign isotropic mechanical properties to material
#ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 25000000, 0.2, 0.0000055)

# define rectangular frame section property
#ret = SapModel.PropFrame.SetRectangle('R1', 'MC25', 0.45, 0.25)

# define rectangular frame section property
ret = SapModel.PropFrame.SetRectangle('B1', 'M25', 0.45, 0.25)
ret = SapModel.PropFrame.SetRebarBeam('B1', 'HYSD500', 'HYSD500', 0.025, 0.025, 804, 804, 402, 402)
ret = SapModel.PropFrame.SetRectangle('C1', 'M25', 0.40, 0.40)
ret = SapModel.PropFrame.SetRebarColumn('C1', 'HYSD500', 'HYSD500', 1, 1, 0.040, 0, 4, 4, "16d", "8d", 0.15, 2, 2, False)
ret = SapModel.PropFrame.SetRectangle('C2', 'MC25', 0.40, 0.40)
ret = SapModel.PropFrame.SetRebarColumn('C2', 'HYSD500', 'HYSD500', 1, 1, 0.040, 0, 4, 4, "16d", "8d", 0.15, 2, 2, False)


# define frame section property modifiers
ModValue = [1, 1, 1, 1, 1, 1, 1, 1]
ret = SapModel.PropFrame.SetModifiers('B1', ModValue)
ret = SapModel.PropFrame.SetModifiers('C1', ModValue)
ret = SapModel.PropFrame.SetModifiers('C2', ModValue)

# add a new mass source and make it the default mass source
#ret = SapModel.SourceMass.SetMassSource("MyMassSource", False, False, True, True, 2, ['1', '2'], [1, 1])

# add frame object by coordinates
FrameName1 = ' '
FrameName2 = ' '
FrameName3 = ' '
FrameName4 = ' '
FrameName5 = ' '
FrameName6 = ' '
FrameName7 = ' '
FrameName8 = ' '
FrameName9 = ' '
FrameName10 = ' '
FrameName11 = ' '
FrameName12 = ' '
FrameName13 = ' '
FrameName14 = ' '
FrameName15 = ' '
FrameName16 = ' '
FrameName17 = ' '
FrameName18 = ' '
FrameName19 = ' '
FrameName20 = ' '
FrameName21 = ' '
FrameName22 = ' '
FrameName23 = ' '
FrameName24 = ' '
FrameName25 = ' '
FrameName26 = ' '
FrameName27 = ' '
FrameName28 = ' '
FrameName29 = ' '
FrameName30 = ' '
FrameName31 = ' '
FrameName32 = ' '
FrameName33 = ' '
FrameName34 = ' '
FrameName35 = ' '
FrameName36 = ' '

[FrameName1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 3.5, FrameName1, 'C1', '1', 'Global')
[FrameName2, ret] = SapModel.FrameObj.AddByCoord(4, 0, 0, 4, 0, 3.5, FrameName2, 'C1', '2', 'Global')
[FrameName3, ret] = SapModel.FrameObj.AddByCoord(8, 0, 0, 8, 0, 3.5, FrameName3, 'C1', '3', 'Global')
[FrameName4, ret] = SapModel.FrameObj.AddByCoord(12, 0, 0, 12, 0, 3.5, FrameName4, 'C1', '4', 'Global')
[FrameName5, ret] = SapModel.FrameObj.AddByCoord(16, 0, 0, 16, 0, 3.5, FrameName5, 'C1', '5', 'Global')
[FrameName6, ret] = SapModel.FrameObj.AddByCoord(0, 0, 3.5, 4, 0, 3.5, FrameName6, 'B1', '6', 'Global')
[FrameName7, ret] = SapModel.FrameObj.AddByCoord(4, 0, 3.5, 8, 0, 3.5, FrameName7, 'B1', '7', 'Global')
[FrameName8, ret] = SapModel.FrameObj.AddByCoord(8, 0, 3.5, 12, 0, 3.5, FrameName8, 'B1', '8', 'Global')
[FrameName9, ret] = SapModel.FrameObj.AddByCoord(12, 0, 3.5, 16, 0, 3.5, FrameName9, 'B1', '9', 'Global')
[FrameName10, ret] = SapModel.FrameObj.AddByCoord(0, 0, 3.5, 0, 0, 7, FrameName10, 'C1', '10', 'Global')
[FrameName11, ret] = SapModel.FrameObj.AddByCoord(4, 0, 3.5, 4, 0, 7, FrameName11, 'C1', '11', 'Global')
[FrameName12, ret] = SapModel.FrameObj.AddByCoord(8, 0, 3.5, 8, 0, 7, FrameName12, 'C1', '12', 'Global')
[FrameName13, ret] = SapModel.FrameObj.AddByCoord(12, 0, 3.5, 12, 0, 7, FrameName13, 'C1', '13', 'Global')
[FrameName14, ret] = SapModel.FrameObj.AddByCoord(16, 0, 3.5, 16, 0, 7, FrameName14, 'C1', '14', 'Global')
[FrameName15, ret] = SapModel.FrameObj.AddByCoord(0, 0, 7, 4, 0, 7, FrameName15, 'B1', '15', 'Global')
[FrameName16, ret] = SapModel.FrameObj.AddByCoord(4, 0, 7, 8, 0, 7, FrameName16, 'B1', '16', 'Global')
[FrameName17, ret] = SapModel.FrameObj.AddByCoord(8, 0, 7, 12, 0, 7, FrameName17, 'B1', '17', 'Global')
[FrameName18, ret] = SapModel.FrameObj.AddByCoord(12, 0, 7, 16, 0, 7, FrameName18, 'B1', '18', 'Global')
[FrameName19, ret] = SapModel.FrameObj.AddByCoord(0, 0, 7, 0, 0, 10.5, FrameName19, 'C2', '19', 'Global')
[FrameName20, ret] = SapModel.FrameObj.AddByCoord(4, 0, 7, 4, 0, 10.5, FrameName20, 'C2', '20', 'Global')
[FrameName21, ret] = SapModel.FrameObj.AddByCoord(8, 0, 7, 8, 0, 10.5, FrameName21, 'C2', '21', 'Global')
[FrameName22, ret] = SapModel.FrameObj.AddByCoord(12, 0, 7, 12, 0, 10.5, FrameName22, 'C2', '22', 'Global')
[FrameName23, ret] = SapModel.FrameObj.AddByCoord(16, 0, 7, 16, 0, 10.5, FrameName23, 'C2', '23', 'Global')
[FrameName24, ret] = SapModel.FrameObj.AddByCoord(0, 0, 10.5, 4, 0, 10.5, FrameName24, 'B1', '24', 'Global')
[FrameName25, ret] = SapModel.FrameObj.AddByCoord(4, 0, 10.5, 8, 0, 10.5, FrameName25, 'B1', '25', 'Global')
[FrameName26, ret] = SapModel.FrameObj.AddByCoord(8, 0, 10.5, 12, 0, 10.5, FrameName26, 'B1', '26', 'Global')
[FrameName27, ret] = SapModel.FrameObj.AddByCoord(12, 0, 10.5, 16, 0, 10.5, FrameName27, 'B1', '27', 'Global')
[FrameName28, ret] = SapModel.FrameObj.AddByCoord(0, 0, 10.5, 0, 0, 14, FrameName28, 'C2', '28', 'Global')
[FrameName29, ret] = SapModel.FrameObj.AddByCoord(4, 0, 10.5, 4, 0, 14, FrameName29, 'C2', '29', 'Global')
[FrameName30, ret] = SapModel.FrameObj.AddByCoord(8, 0, 10.5, 8, 0, 14, FrameName30, 'C2', '30', 'Global')
[FrameName31, ret] = SapModel.FrameObj.AddByCoord(12, 0, 10.5, 12, 0, 14, FrameName31, 'C2', '31', 'Global')
[FrameName32, ret] = SapModel.FrameObj.AddByCoord(16, 0, 10.5, 16, 0, 14, FrameName32, 'C2', '32', 'Global')
[FrameName33, ret] = SapModel.FrameObj.AddByCoord(0, 0, 14, 4, 0, 14, FrameName33, 'B1', '33', 'Global')
[FrameName34, ret] = SapModel.FrameObj.AddByCoord(4, 0, 14, 8, 0, 14, FrameName34, 'B1', '34', 'Global')
[FrameName35, ret] = SapModel.FrameObj.AddByCoord(8, 0, 14, 12, 0, 14, FrameName35, 'B1', '35', 'Global')
[FrameName36, ret] = SapModel.FrameObj.AddByCoord(12, 0, 14, 16, 0, 14, FrameName36, 'B1', '36', 'Global')

# assign point object restraint at base
PointName1 = ' '
PointName2 = ' '
Restraint1 = [True, True, True, True, True, True]
Restraint2 = [ False, True, False, True, False, True]

# Fixed restrained value assign
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint1)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint1)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint1)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName4, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint1)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName5, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint1)
# refresh view, update (initialize) zoom
ret = SapModel.View.RefreshView(0, False)

# add load patterns
LTYPE_1 = 1
LTYPE_2 = 2
LTYPE_3 = 3
LTYPE_5 = 5
LTYPE_8 = 8

ret = SapModel.LoadPatterns.Add('1', LTYPE_2, 1, True)
ret = SapModel.LoadPatterns.Add('2', LTYPE_8, 0, True)
ret = SapModel.LoadPatterns.Add('Eq', LTYPE_8, 0, True)

# assign loading for load pattern 2 at floor levels
ret = SapModel.FrameObj.SetLoadDistributed(FrameName6, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName7, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName8, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName9, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName15, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName16, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName17, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName18, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName24, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName25, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName26, '2', 1, 10, 0, 1, 49.45, 49.55)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName27, '2', 1, 10, 0, 1, 49.45, 49.55)

# assign loading for load pattern 2 at the roof level
ret = SapModel.FrameObj.SetLoadDistributed(FrameName33, '2', 1, 10, 0, 1, 34.72, 34.72)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName34, '2', 1, 10, 0, 1, 34.72, 34.72)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName35, '2', 1, 10, 0, 1, 34.72, 34.72)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName36, '2', 1, 10, 0, 1, 34.72, 34.72)

# assign loading for load pattern Eq
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName28, PointName1, PointName2)
PointLoadValue = [43.6, 0, 0, 0, 0, 0]
ret = SapModel.PointObj.SetLoadForce(PointName2, 'Eq', PointLoadValue)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName19, PointName1, PointName2)
PointLoadValue = [46.08, 0, 0, 0, 0, 0]
ret = SapModel.PointObj.SetLoadForce(PointName2, 'Eq', PointLoadValue)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName10, PointName1, PointName2)
PointLoadValue = [20.49, 0, 0, 0, 0, 0]
ret = SapModel.PointObj.SetLoadForce(PointName2, 'Eq', PointLoadValue)
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
PointLoadValue = [5.13, 0, 0, 0, 0, 0]
ret = SapModel.PointObj.SetLoadForce(PointName2, 'Eq', PointLoadValue)
