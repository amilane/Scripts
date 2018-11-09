# -*- coding: utf-8 -*-

import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")
import string
from operator import itemgetter
from itertools import groupby

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

NONELIST = [None, '', ' ']
literals = string.ascii_uppercase
projInfo = doc.ProjectInformation
tag1Value = projInfo.LookupParameter('TagCode1').AsString()
tag2Value = projInfo.LookupParameter('TagCode2').AsString()

def filterWithCode(l):
	return list(filter(lambda x: x.LookupParameter('TagCode3').AsString() != None and x.LookupParameter('TagCode4').AsString() != None, l))

def groupByFam(l):
	# группировка по "Семейство и типоразмер"
	sList = sorted(l, key = lambda e: e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString())
	gList = [[x for x in g] for k,g in groupby(sList, lambda e: e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString())]
	return gList

def groupBySize(gList):
	# группировка по размерам
	gSize = []
	for g in gList:
		sG = sorted(g, key = lambda e: e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString())
		gSG = [[x for x in g] for k,g in groupby(sG, lambda e: e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString())]
		gSize += gSG
	return gSize

def groupByTag3_4(gList):
	# группировка по размерам TagCode3
	sList = sorted(gList, key = lambda g: g[0].LookupParameter('TagCode3').AsString()+g[0].LookupParameter('TagCode4').AsString())
	gList2 = [[x for x in g] for k,g in groupby(sList, lambda g: g[0].LookupParameter('TagCode3').AsString()+g[0].LookupParameter('TagCode4').AsString())]
	return gList2

def getUsedTag5(gList):
	used = list(set([int(e.LookupParameter('TagCode5').AsString()) for g in gList for e in g if e.LookupParameter('TagCode5').AsString() not in NONELIST]))
	if used:
		return used
	else:
		return [0]

def groupBySys(l):
	sList = sorted(l, key = lambda e: e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString())
	gList = [[x for x in g] for k,g in groupby(sList, lambda e: e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString())]
	return gList

def groupDuctsBySizeAndThi(ducts):
	ductsWithKeys = [[e, '-'.join([e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString(),\
		e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString(),\
		e.LookupParameter('AG_Thickness').AsValueString()])] for e in ducts]
	sList = sorted(ductsWithKeys, key = lambda e: e[1])
	gList = [[x[0] for x in g] for k,g in groupby(sList, lambda e: e[1])]
	return gList

def findTag5Value(g):
	taggedElem = filter(lambda e: e.LookupParameter('TagCode5').AsString() not in NONELIST, g)
	if taggedElem:
		tag5Value = taggedElem[0].LookupParameter('TagCode5').AsString()
	else:
			i = 1
			while i < max(USEDCode5):
				if i not in USEDCode5:
					tv = i
					USEDCode5.append(tv)
					break
				i += 1
			else:
				tv = max(USEDCode5) + 1
				USEDCode5.append(tv)
			tag5Value = str(tv)
	return tag5Value

def groupFitBySizeAndPartType(fittings):
	fitWithKeys = []
	lParNames = ('Length', 'Длина воздуховода', 'L')
	for e in fittings:
		a1 = str(e.MEPModel.PartType)
		a2 = e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
		try:
			for p in lParNames:
				a3 = e.LookupParameter(p).AsString()
		except: a3 = ''
		a4 = e.LookupParameter('AG_Thickness').AsValueString()
		fitWithKeys.append([e, '-'.join([a1, a2, a3, a4])])
	sList = sorted(fitWithKeys, key = lambda e: e[1])
	gList = [[x[0] for x in g] for k,g in groupby(sList, lambda e: e[1])]
	return gList

# воздухораспределители
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# арматура
ductAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# арматура труб
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
# оборудование
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()

# воздуховоды
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
# гибкий воздуховод
flexDuct = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
# соед детали воздуховодов
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()

gTerminal = groupBySize(groupByFam(filterWithCode(terminal)))
gDuctAcc = groupBySize(groupByFam(filterWithCode(ductAccessory)))
gPipeAcc = groupBySize(groupByFam(filterWithCode(pipeAccessory)))
gEquip = groupByFam(filterWithCode(equipment))

allGroups = gTerminal + gDuctAcc + gPipeAcc + gEquip
allGroupsTag3 = groupByTag3_4(allGroups)

allDuctElems = list(ducts) + list(flexDuct) + list(fitings)
groupDuctElemsBySys = groupBySys(allDuctElems)


t = Transaction(doc, 'Setparameter')
t.Start()
# арматура труб, арматура воздуховодов, воздухораспределители, оборудование
for group in allGroupsTag3:
	USEDCode5 = getUsedTag5(group)
	for g in group:
		# нумерация 0001, 0002, 0003
		tag5Value = findTag5Value(g)

		# уже задействованные буквы в группах
		USEDCode6 = list(set([e.LookupParameter('TagCode6').AsString() for e in g if e.LookupParameter('TagCode6').AsString() not in NONELIST]))
		
		k = 0
		for e in g:
			# нумерация A, B, C, D...
			tag5 = e.LookupParameter('TagCode5')
			tag5Value = tag5Value.zfill(4)
			tag5.Set(tag5Value)
			
			tag6Value = ''
			tag6 = e.LookupParameter('TagCode6')
			if USEDCode6:
				if tag6.AsString() not in USEDCode6:
					n = 0
					while n < literals.index(max(USEDCode6)):
						if literals[n] not in USEDCode6:
							tag6Value = literals[n]
							USEDCode6.append(tag6Value)
							break
						n += 1
					else:
						n += 1
						tag6Value = literals[n]
						USEDCode6.append(tag6Value)
					tag6.Set(tag6Value)
			else:
				tag6Value = literals[k]
				tag6.Set(tag6Value)
				k += 1
			
			tag3Value = e.LookupParameter('TagCode3').AsString()
			tag4Value = e.LookupParameter('TagCode4').AsString()
			tag5Value = e.LookupParameter('TagCode5').AsString()
			tag6Value = e.LookupParameter('TagCode6').AsString()
			
			tagsToStr = []
			for i in [tag1Value, tag2Value, tag3Value, tag4Value, tag5Value, tag6Value]:
				if i != None:
					tagsToStr.append(i)
				else:
					tagsToStr.append('')

			tagAllValue = '-'.join(tagsToStr[:5]) + tagsToStr[-1]
			tagAll = e.LookupParameter('TAG')
			tagAll.Set(tagAllValue)


# воздуховоды и фасонка
for sysgroup in groupDuctElemsBySys:

	USEDCode5 = list(set([int(e.LookupParameter('TagCode5').AsString()) for e in sysgroup if e.LookupParameter('TagCode5').AsString() not in NONELIST]))
	if USEDCode5 == []:
		USEDCode5 = [0]

	ducts = filter(lambda e: e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves), sysgroup)
	ductsBySizeAndThi = groupDuctsBySizeAndThi(ducts)

	for g in ductsBySizeAndThi:
		tag5Value = findTag5Value(g)
		for e in g:
				tag5 = e.LookupParameter('TagCode5')
				tagAll = e.LookupParameter('TAG')
				tag5Value = tag5Value.zfill(3)
				tag5.Set(tag5Value)
				sp = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
				tagAllValue = '/'.join([tag1Value, sp, tag5Value])
				tagAll.Set(tagAllValue)

	fittings = filter(lambda e: e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting), sysgroup)
	fitBySize = groupFitBySizeAndPartType(fittings)

	for g in fitBySize:
		tag5Value = findTag5Value(g)
		for e in g:
				tag5 = e.LookupParameter('TagCode5')
				tagAll = e.LookupParameter('TAG')
				tag5Value = tag5Value.zfill(3)
				tag5.Set(tag5Value)
				tagAllValue = '/'.join([tag1Value, sp, tag5Value])
				tagAll.Set(tagAllValue)


t.Commit()

MessageBox.Show("OK", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

