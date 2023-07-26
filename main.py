# -*- coding: utf-8 -*-

from openpyxl import load_workbook, Workbook

from localVariables import skillList, raceClassDict, guildList, spellList

source_folder = "d:\Work\Coding\\battleboard_data_extraction\\battleboard-data-extraction\source_files\\"

fileList = ["Wulfric_baneguard_current_v_2021.1.xlsm",
            "emilie_oct_23.xlsm",
            ]

# test_workbook = load_workbook(source_folder + "Wulfric_baneguard_current_v_2021.1.xlsm")

# print(test_workbook.sheetnames)
# ['Instructions', 'Adventure Record', 'The Character', 'Magic', 'Power', 'Armour', 'Battleboard', 'Notes', 'Base', 'Variables', 'Tables', 'STD HQA']

# skill name maps to dict with 2 keys, 1 to track total ranks, 1 to track number of files that formed the total
# should look like
# skillMap = {'Bonebreak' : {'ranks' : 4,
#                            'characters' : 2},
#             'dodge'     : {'ranks' : 1,
#                            'characters' : 1}
#             } 

skillMap = {}

# spell name maps to dict with 4 leys, each to track the number of times bought by a character of a given archetype
# should look like
# spellMap = {'blindness' : {'warrior' : 0,
#                            'scout' : 2,
#                            'acolyte' : 1,
#                            'mage' : 4},
#             'lightning bolt' : {'warrior' : 0,
#                                 'scout' : 2,
#                                 'acolyte' : 1,
#                                 'mage' : 4},
#             } 

spellMap = {}

for file in fileList:

    currentWorkbook = load_workbook(source_folder + file)
    skillSheet = currentWorkbook['The Character']

    characterClass = raceClassDict[skillSheet.cell(row=2,column=2).value]['class']
    characterRace = raceClassDict[skillSheet.cell(row=2,column=2).value]['race']
    primaryGuild = guildList[(skillSheet.cell(row=3,column=2).value - 1)]

    # openpyxl starts at 1 for row and column
    for i in range(14,397):
        
        # allow for offset between row number on sheet and position in list
        skillName = skillList[(i-14)]
        rankValue = skillSheet.cell(row=i,column=3).value 

        if rankValue != None:

            if skillName not in skillMap:
                skillMap[skillName] = {'ranks' : rankValue,
                                       'characters' : 1}
            else:
                skillMap[skillName]['ranks'] = skillMap[skillName]['ranks'] + rankValue
                skillMap[skillName]['characters'] = skillMap[skillName]['characters'] + 1 

    spellSheet = currentWorkbook['Magic']

    for i in range(12,427):

        spellBought = spellSheet.cell(row=i,column=4).value
        # print(spellBought)
        # break
        spellName = spellList[(i-12)]

        if spellBought == None or spellBought == 0:
            continue
        else:
            if spellName not in spellMap:
                spellMap[spellName] = {'Warrior' : 0,
                                       'Scout' : 0,
                                       'Acolyte' : 0,
                                       'Mage' : 0}
            
            spellMap[spellName][characterClass] += 1

        

result = Workbook()
skillSheet = result.active
skillSheet.title = "Skill summary"

skillSheet.cell(row=1,column=1).value = 'Skill Name'
skillSheet.cell(row=1,column=2).value = 'Average Ranks'
skillSheet.cell(row=1,column=3).value = 'Character Count'

rowNum = 2
for skill in skillList:
    skillSheet.cell(row=rowNum, column=1).value = skill

    if skill in skillMap:
        numCharacters = skillMap[skill]['characters']
        skillSheet.cell(row=rowNum, column=2).value = (skillMap[skill]['ranks'] / numCharacters)
        skillSheet.cell(row=rowNum, column=3).value = numCharacters
    else:
        skillSheet.cell(row=rowNum, column=2).value = 0
        skillSheet.cell(row=rowNum, column=3).value = 0

    rowNum += 1

spellSheet = result.create_sheet("Spells Summary")

spellSheet.cell(row=1,column=1).value = 'Spell Name'
spellSheet.cell(row=1,column=2).value = 'Times bought Mage'
spellSheet.cell(row=1,column=3).value = 'Times bought Acolyte'
spellSheet.cell(row=1,column=4).value = 'Times bought Scout'
spellSheet.cell(row=1,column=5).value = 'Times bought Warrior'

rowNum = 2
for spell in spellList:
    spellSheet.cell(row=rowNum, column=1).value = spell

    if spell in spellMap:
        spellSheet.cell(row=rowNum, column=2).value = spellMap[spell]['Mage']
        spellSheet.cell(row=rowNum, column=3).value = spellMap[spell]['Acolyte']
        spellSheet.cell(row=rowNum, column=4).value = spellMap[spell]['Scout']
        spellSheet.cell(row=rowNum, column=5).value = spellMap[spell]['Warrior']
    else:
        spellSheet.cell(row=rowNum, column=2).value = 0
        spellSheet.cell(row=rowNum, column=3).value = 0
        spellSheet.cell(row=rowNum, column=4).value = 0
        spellSheet.cell(row=rowNum, column=5).value = 0

    rowNum += 1

result.save(source_folder + "results.xlsx")

# print(len(skillList))
# print(skillMap)
# print(spellList)
# print(spellMap)

