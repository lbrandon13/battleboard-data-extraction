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

        

# result = Workbook()
# sheet = result.active
# sheet.title = "Skill summary"

# sheet.cell(row=1,column=1).value = 'Skill Name'
# sheet.cell(row=1,column=2).value = 'Average Ranks'
# sheet.cell(row=1,column=3).value = 'Character Count'

# rowNum = 2
# for skill in skillList:
#     sheet.cell(row=rowNum, column=1).value = skill

#     if skill in skillMap:
#         numCharacters = skillMap[skill]['characters']
#         sheet.cell(row=rowNum, column=2).value = (skillMap[skill]['ranks'] / numCharacters)
#         sheet.cell(row=rowNum, column=3).value = numCharacters
#     else:
#         sheet.cell(row=rowNum, column=2).value = 0
#         sheet.cell(row=rowNum, column=3).value = 0

#     rowNum += 1


# result.save(source_folder + "results.xlsx")

# print(len(skillList))
# print(skillMap)
# print(spellList)
print(spellMap)

