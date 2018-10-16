# -*- coding: utf-8 -*-
# заполнение непользовательских параметров в элементах (Система, ТолщинаУгол, Размер, Количество, ЕдИзм, Уровень, Код категории, Наименование (генерит для воздуховодов и фасонки предварительно))
# продумать алгоритм расчета толщины гофрированных труб

import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

antismoke = ["ВПВ", "ППВ", "ДУ", "КДУ", "ПВ"]
# функция определения настоящего уровня элемента


def tByConnector(sysName, con, typeIsol):
	shape = con.Shape
	if shape == ConnectorProfileType.Rectangular:
		h = con.Height
		w = con.Width
		maxS = max(h, w)
		maxMM = UnitUtils.ConvertFromInternalUnits(maxS, DisplayUnitType.DUT_MILLIMETERS)
		if typeIsol != None and "EI" in typeIsol:
			if maxMM <= 2000.0:
				t = 0.9
			else:
				t = 1.2
		else:
			if sysName != None and any(sysName.__contains__(i) for i in antismoke):
				if maxMM <= 2000.0:
					t = 0.9
				else:
					t = 1.2
			else:
				if maxMM <= 250.0:
					t = 0.5
				elif maxMM <= 1000:
					t = 0.7
				else:
					t = 0.9
	elif shape == ConnectorProfileType.Round:
		d = con.Radius * 2
		dMM = UnitUtils.ConvertFromInternalUnits(d, DisplayUnitType.DUT_MILLIMETERS)
		if typeIsol != None and "EI" in typeIsol:
			if dMM <= 2000.0:
				t = 0.9
			else:
				t = 1.2
		else:
			if sysName != None and any(sysName.__contains__(i) for i in antismoke):
				if dMM <= 800.0:
					t = 0.9
				elif dMM <= 1250.0:
					t = 1.0
				elif dMM <= 1600.0:
					t = 1.2
				else:
					t = 1.4
			else:
				if dMM <= 200.0:
					t = 0.5
				elif dMM <= 450.0:
					t = 0.6
				elif dMM <= 800.0:
					t = 0.7
				elif dMM <= 1250.0:
					t = 1.0
				elif dMM <= 1600.0:
					t = 1.2
				else:
					t = 1.4
	return t
# заполнить параметр Толщина Угол для воздуховодов
def setThiDucts(e):
	sysName = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
	spcThiAngle = e.LookupParameter("AG_Thickness")
	if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
		cc = [i for i in e.ConnectorManager.Connectors]
		con = cc[0]
		typeIsol = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
		thi = tByConnector(sysName, con, typeIsol)
	elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
		thi = 0.15
	spcThiAngle.Set(thi)

# заполнить параметр Толщина Угол для фитингов
def setThiItems(e):
	sysName = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
	spcThi = e.LookupParameter("AG_Thickness")
	thi = 0
	if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting):
		cc = [i for i in e.MEPModel.ConnectorManager.Connectors]
		maxConSquare = 0.0
		for con in cc:
			if con.Shape == ConnectorProfileType.Rectangular:
				conSquare = con.Height * con.Width
			elif con.Shape == ConnectorProfileType.Round:
				conSquare = math.pi * con.Radius ** 2
			if conSquare > maxConSquare:
				outCon = con
		typeIsol = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
		thi = tByConnector(sysName, outCon, typeIsol)
	elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
		cc = [i for i in e.MEPModel.ConnectorManager.Connectors]
		
		for con in cc:
			allRefs = [i for i in con.AllRefs]
			for i in allRefs:
				own = i.Owner
				if own.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves):
					thi = own.LookupParameter("AG_Thickness").AsString()
					break
	else:
		thi = 0
	spcThi.Set(thi)

# воздуховоды
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
# соед детали воздуховодов
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
# уровни
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
levelsSort = sorted(levels, key = lambda i: i.Elevation)


t = Transaction(doc, "SetParameters")
t.Start()

# фитинги
if fitings:
	for e in fitings:
		setThiItems(e)

# воздуховоды
if ducts:
	for e in ducts:
		setThiDucts(e)

t.Commit()

MessageBox.Show("ОК", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

