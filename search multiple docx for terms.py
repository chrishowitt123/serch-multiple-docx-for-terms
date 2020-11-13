# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 14:43:49 2020
@author: chris
"""
import pandas as pd
import tkinter
from tkinter import filedialog
from iteration_utilities import deepflatten
from collections import Counter
import os
import docx2txt
from nltk.tokenize import PunktSentenceTokenizer
from nltk.tokenize import RegexpTokenizer
import glob
from pathlib import Path    


print( "\n")  
print("Welcome to DOCX_RepStrings!")
print( "\n")  
print("One file or multiple files?")
print( "\n") 

answer = input("One file or multiple files? 1 / M: ")
print( "\n")

if answer == str(1):

    root = tkinter.Tk()
    root.wm_withdraw() # this completely hides the root window
    filename = filedialog.askopenfilename()
    
    fn = filename.split('.')[0] #seperate filename from extention for f-string output files
    
    text = docx2txt.process(filename)
    sent_tokenizer = PunktSentenceTokenizer(text)
    sents = sent_tokenizer.tokenize(text)
    
    allTextList = []
    
    for sent in sents:
        s = sent.split("\n")
        allTextList.append(s)
    
    allTextFlat = list(deepflatten(allTextList, depth=1))
    allTextFlat = [ x.replace('\t', '') for x in allTextFlat ]
    allTextFlat = [x for x in allTextFlat if x.strip()] 
    
    repeaters  = Counter(allTextFlat)
    
    results = pd.DataFrame.from_dict(repeaters, orient='index')
    results.columns = ['No. of instances']  
    results = results[results['No. of instances']>1]
    results.sort_values(by=['No. of instances'],ascending=False, inplace=True)
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    
    results.reset_index(inplace=True)
    results = results.rename(columns = {'index':'String'})
    
    strings = results['String'].tolist()
    insts = results['No. of instances'].tolist()
    
    dml = []
    
    for string, inst in zip(strings, insts):
        string_word_count = len(string.split(' '))
        dup_material = string_word_count * inst - string_word_count
        dml.append(dup_material)
    
        word_tokenizer = RegexpTokenizer(r'\w+')
        tokens = word_tokenizer.tokenize(text)
        
        rep_words = sum(dml)
        doc_word_count = len(tokens)
        perc_rep_mat = (rep_words / doc_word_count) * 100
        rep_mat_statment = f'     {round(perc_rep_mat, 2)} % repeated material.'
        results['%'] = pd.Series(rep_mat_statment)
        
        
        
        
    os.path.dirname(os.path.abspath(filename))
    results.to_excel(filename + '_DOCX_RepStrings.xlsx')
        
    print("Results!")
    print( "\n")
    print(results[['String', 'No. of instances']])
    print( "\n")
    print(rep_mat_statment)  
    print( "\n")
    print("Your DOCX_RepStrings report was successfully created!")
        
else:
    
    root = tkinter.Tk()
    root.wm_withdraw() # this completely hides the root window
    folder = filedialog.askdirectory()
    orignalExt = '\*.docx'
    orignalFiles = folder + orignalExt
    
    sourceFiles = []
    for filepath in glob.iglob(orignalFiles):
        sourceFiles.append(filepath)
        
        for filename in sourceFiles:
            
    
            fn = filename.split('.')[0] #seperate filename from extention for f-string output files
                
            text = docx2txt.process(filename)
            sent_tokenizer = PunktSentenceTokenizer(text)
            sents = sent_tokenizer.tokenize(text)
            
        allTextList = []
        
        for sent in sents:
            s = sent.split("\n")
            allTextList.append(s)
        
        allTextFlat = list(deepflatten(allTextList, depth=1))
        allTextFlat = [ x.replace('\t', '') for x in allTextFlat ]
        allTextFlat = [x for x in allTextFlat if x.strip()] 
        
        repeaters  = Counter(allTextFlat)
        
        results = pd.DataFrame.from_dict(repeaters, orient='index')
        results.columns = ['No. of instances']  
        results = results[results['No. of instances']>1]
        results.sort_values(by=['No. of instances'],ascending=False, inplace=True)
        pd.set_option("display.max_rows", None, "display.max_columns", None)
        
        results.reset_index(inplace=True)
        results = results.rename(columns = {'index':'String'})
        
        strings = results['String'].tolist()
        insts = results['No. of instances'].tolist()
        
        dml = []
        
        for string, inst in zip(strings, insts):
            string_word_count = len(string.split(' '))
            dup_material = string_word_count * inst - string_word_count
            dml.append(dup_material)
        
        word_tokenizer = RegexpTokenizer(r'\w+')
        tokens = word_tokenizer.tokenize(text)
        
        rep_words = sum(dml)
        doc_word_count = len(tokens)
        perc_rep_mat = (rep_words / doc_word_count) * 100
        rep_mat_statment = f'     {round(perc_rep_mat, 2)} % repeated material.'
        results['%'] = pd.Series(rep_mat_statment)
        
        
        
        results.to_excel(filename + '_DOCX_RepStrings.xlsx')
 

        f=open(f"{folder}/rep_mat_%.txt", "a+")
        f.write(f'{Path(filename).name} \t \t \t \t ---> \t \t \t {rep_mat_statment}\n')
        
        
        print(f'{Path(filename).name[:20]}.. \t  ---> {rep_mat_statment}')
        
        
        
        print( "\n")
        
    f.close()
    print("Your DOCX_RepStrings reports were successfully created!")
