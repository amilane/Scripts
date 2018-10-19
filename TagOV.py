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

def groupByTag4(gList):
	# группировка по размерам TagCode4
	gTag = []
	for g in gList:
		sG = sorted(g, key = lambda e: e.LookupParameter('TagCode4').AsString())
		gSG = [[x for x in g] for k,g in groupby(sG, lambda e: e.LookupParameter('TagCode4').AsString())]
		gTag += gSG
	return gTag
'''
# не нужно, т.к. ADSK_Предел огнестойкости - параметр типа
def groupByFire(gList):
	#группировка по огнестойкости для арматуры воздуховодов
	gFire = []
	for g in gList:
		sG = sorted(g, key = lambda e: e.LookupParameter('ADSK_Предел огнестойкости').AsString())
		gSG = [[x for x in g] for k,g in groupby(sG, lambda e: e.LookupParameter('ADSK_Предел огнестойкости').AsString())]
		gFire += gSG
	return gFire
'''
def getUsedTag5(gList):
	used = list(set([int(e.LookupParameter('TagCode5').AsString()) for g in gList for e in g if e.LookupParameter('TagCode5').AsString() not in NONELIST]))
	if used:
		return used
	else:
		return [0]


# воздухораспределители
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# арматура
ductAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# арматура труб
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
# оборудование
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()

gTerminal = groupBySize(groupByFam(terminal))
gDuctAcc = groupBySize(groupByFam(ductAccessory))
gPipeAcc = groupBySize(groupByFam(pipeAccessory))
gEquip = groupByFam(equipment)

allGroups = gTerminal + gDuctAcc + gPipeAcc + gEquip



USEDCode5 = getUsedTag5(gTerminal)

t = Transaction(doc, 'Setparameter')
t.Start()
for g in allGroups:
	# нумерация 0001, 0002, 0003
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

		tagAllValue = '-'.join(tagsToStr)
		tagAll = e.LookupParameter('TAG')
		tagAll.Set(tagAllValue)
t.Commit()

MessageBox.Show(str("OK"), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

