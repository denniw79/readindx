# -*- coding: utf-8 -*-
"""
Created on Wed Mar 10 20:08:28 2021
Spyder Editor

DWERCreadabilityCalcultor

- Input: xslx-table with given Text per line in column b *** to build xlsx-table from txt-files use DWExlsxGenerator ***
- Calculation of readability indices [columns d-g]: Flesch-Kincaid-Grade-Level [FKGL], Flesch-Reading-Ease (Deutsch) [FRE], Gunning-Fog-Index [GFI], Automated readability index [ARI]
- Calculation of following parameters per text/line [columens h-n]: No. of sentences, subsentences, words, syllables, nouns, verbs, adjectives
- calculation of word frequences based on SUBTLEX-DE [column o]

Installation of Treetagger and the Treetagger-Wrapper necessary for count of specific word types and word frequency matching:
    https://treetaggerwrapper.readthedocs.io/en/latest/

@author: Dennis[DWE]Winter-Extra
"""

import pandas as pd
import textstat as txtst
import treetaggerwrapper
# Progressbar is in file in same directory, not installed separately
import progressbar
from collections import Counter

# Fuction calculating Flesch-Reading-Ease-Index
def fre_deutsch(asl, asw):
    return 180 - asl - (58.5 * asw)
    
# Function matching word frequencies
def findSUBTLEX(word):
    global foundwordcount
    searchDF = frequencies.loc[frequencies["Word"] == word]["SUBTLEX"]
    if len(searchDF)>0:
        foundwordcount +=1
        return searchDF.iloc[0]
    else: 
        return 0

file_name =  'bundledTexts.xlsx'
output_file_name = 'DWERCoutput.xlsx'
sheet = 0


# Create treetagger for German tagging
tagger = treetaggerwrapper.TreeTagger(TAGLANG='de')

# language setting for index-calculations
txtst.set_lang('de_DE')

output = []
foundwordcount = 0

# Read data from Excel
df = pd.read_excel(io=file_name, sheet_name=sheet)

frequencies = pd.read_excel(io='SUBTLEX-DE cleaned with Google00 frequencies.xlsx', sheet_name=0)


# Determine the length of the file (for progress bar)
totalrows = len(df.index)

# Iterate over all rows in the dataframe 
for index, row in df.iterrows():
    
    # Update progress bar. Not really necessary, but nice for longer operations
    progressbar.printProgressBar(index+1, totalrows, prefix = "Fortschritt: ", suffix = "Erledigt", length = 50)    
    targettxt = row[1].strip()
    # Calculate different readability indices
    fleschk = txtst.flesch_kincaid_grade(targettxt)
    gunningf = txtst.gunning_fog(targettxt)
    ari = txtst.automated_readability_index(targettxt)
    
    
    #FRE berechnen    
    asl = txtst.lexicon_count(targettxt) / txtst.sentence_count(targettxt)
    asw = txtst.syllable_count(targettxt)/txtst.lexicon_count(targettxt)
    fre = fre_deutsch(asl, asw)
    
    # Tag Wordtypes using treetagger 
    tags = treetaggerwrapper.make_tags(tagger.tag_text(targettxt))
    nouncount = 0
    verbcount = 0 
    adjectivecount = 0
    frequencysum = 0
    frequencyavg = 0
    foundwordcount = 0
    
    # Iterate over treetagger results and count Verbs, Nouns and Adjectives. 
    for tag in tags:
        if isinstance(tag, treetaggerwrapper.Tag):
            wordtype = tag.pos[0]
            if wordtype =='V' :
                verbcount +=1
                frequencysum += findSUBTLEX(tag.word)
            
            elif wordtype == 'N':
                nouncount +=1
                frequencysum += findSUBTLEX(tag.word)
            elif (wordtype =='A') and tag.pos[0:3] == 'ADJ' :
                    adjectivecount +=1
                    frequencysum += findSUBTLEX(tag.word)
    
    
    if foundwordcount>0:     
        frequencyavg = frequencysum / foundwordcount
        
    # Count punctuation marks for subsentences. 
    # Source: https://stackoverflow.com/questions/6969268/counting-letters-numbers-and-punctuation-in-a-string/14229674
    counts = Counter(targettxt)
    punctuation_count = counts['.'] + counts[','] + counts[';'] + counts['!'] + counts[':'] + counts['?']
    
    # Add results to list 
    output.append([row[0], targettxt, row[2], fre, fleschk, gunningf, ari, 
                       txtst.sentence_count(targettxt), punctuation_count, txtst.lexicon_count(targettxt), 
                       txtst.syllable_count(targettxt), nouncount, verbcount, adjectivecount, frequencyavg])

    # Liste in Dataframe
    outputdf = pd.DataFrame(output, columns=['Satznummer', 'Satz', 'Version', 'FRE_de','FKGL', 'GFI',
                                         'ARI', '#Sätze', '#Teilsätze', '#Wörter', 'Silben', 
                                         'Nomen', 'Verben', 'Adjektive', 'Durchschn. Häufigkeit'])

# Write dataframe to Excel
outputdf.to_excel(output_file_name, index = False)

