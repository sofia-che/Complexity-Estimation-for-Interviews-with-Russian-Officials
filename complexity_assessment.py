import pandas as pd
df = pd.read_excel('corpus_annot_all.xlsx')

def calculate_metrics(group):
    words = [str(i) for i in group['token'].tolist()]
    lemmas = [str(i) for i in group['lemma'].tolist()]
    upos = [str(i) for i in group['upos'].tolist()]
    feats = [str(i) for i in group['feats'].tolist()]
    dep_rel = [str(i) for i in group['dep_rel'].tolist()]

    def NVR(pos):
        n = 0
        v = 0
        for i in pos:
            if i == 'NOUN' or i == 'PROPN':
                n += 1
            elif i == 'VERB' or i == 'AUX':
                v += 1
        if v != 0:
            return n / v
        else:
            return 0

    def Autosem_pr(pos):
        count = 0
        for i in pos:
            if i == 'ADJ' or i == 'ADV' or i == 'NOUN' or i == 'NUM' or i == 'PROPN' or i == 'VERB':
                count += 1
        return count / len(pos)

    def Func_word_pr(pos):
        count = 0
        for i in pos:
            if i == 'ADP' or i == 'AUX' or i == 'CCONJ' or i == 'PART' or i == 'SCONJ':
                count += 1
        return count / len(pos)

    def Verb_pr(pos):
        count = 0
        for i in pos:
            if i == 'VERB' or i == 'AUX':
                count += 1
        return count / len(pos)

    def Noun_pr(pos):
        count = 0
        for i in pos:
            if i == 'NOUN' or i == 'PROPN':
                count += 1
        return count / len(pos)

    def Adj_pr(pos):
        count = 0
        for i in pos:
            if i == 'ADJ':
                count += 1
        return count / len(pos)

    def Pron_pr(pos):
        count = 0
        for i in pos:
            if i == 'DET' or i == 'PRON':
                count += 1
        return count / len(pos)

    def Sconj_pr(pos):
        count = 0
        for i in pos:
            if i == 'SCONJ':
                count += 1
        return count / len(pos)

    def Cconj_pr(pos):
        count = 0
        for i in pos:
            if i == 'CCONJ':
                count += 1
        return count / len(pos)

    def Word_form(lemmas):
        count = 0
        for i in lemmas:
            if i.endswith(
                    tuple(
                        ['ция', 'ние', 'вие', 'тие', 'ист', 'изм', 'ура', 'ище', 'ство', 'ость', 'овка', 'атор', 'итор',
                         'тель', 'льный', 'овать'])):
                count += 1
        return count / len(lemmas)

    def Gen_pr(feats):
        count = 0
        for i in feats:
            if 'Case=Gen' in str(i):
                count += 1
        return count / len(feats)

    def Ablt_pr(feats):
        count = 0
        for i in feats:
            if 'Case=Ins' in str(i):
                count += 1
        return count / len(feats)

    def dat(feats):
        count = 0
        for i in feats:
            if 'Case=Dat' in str(i):
                count += 1
        return count / len(feats)

    def nomn(feats):
        count = 0
        for i in feats:
            if 'Case=Nom' in str(i):
                count += 1
        return count / len(feats)

    def loct(feats):
        count = 0
        for i in feats:
            if 'Case=Loc' in str(i):
                count += 1
        return count / len(feats)

    def Neut_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'NOUN':
                if 'Gender=Neut' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def Inan_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'NOUN':
                if 'Animacy=Inan' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def firstp_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Person=1' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def thirdp_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Person=3' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def pres_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Tense=Pres' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def futr_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Tense=Fut' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def past_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Tense=Past' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def impf_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Aspect=Imp' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def pf_pr(pos, feats):
        count = 0
        for i in range(len(pos)):
            if str(pos[i]) == 'VERB' or str(pos[i]) == 'AUX':
                if 'Aspect=Perf' in str(feats[i]):
                    count += 1
        return count / len(feats)

    def pssv_prtf_pr(feats):
        count = 0
        for i in feats:
            if 'VerbForm=Part' in str(i) and 'Voice=Pass' in str(i) and 'Variant=Short' not in str(i):
                count += 1
        return count / len(feats)

    def pssv_prts_pr(feats):
        count = 0
        for i in feats:
            if 'VerbForm=Part' in str(i) and 'Voice=Pass' in str(i) and 'Variant=Short' in str(i):
                count += 1
        return count / len(feats)

    def sya_forms(words, pos):
        count = 0
        for i in range(len(words)):
            if str(pos[i]) == 'VERB' and str(words[i]).endswith('ся'):
                count += 1
        return count / len(words)

    def yavl_pr(lemmas):
        count = 0
        for i in lemmas:
            if str(i).lower() == 'являться':
                count += 1
        return count / len(words)

    with open('Abstract.txt', encoding='utf-8') as file:
        lines = file.readlines()
        Abstract = [line.rstrip() for line in lines]

    def abstr_pr(lemmas):
        count = 0
        for i in lemmas:
            if str(i).lower() in Abstract:
                count += 1
        return count / len(words)

    with open('Deont.txt', encoding='utf8') as file:
        lines = file.readlines()
        Deont = [line.rstrip() for line in lines]

    def deont_pr(lemmas):
        count = 0
        newlemmas = []
        for i in lemmas:
            newi = str(i)
            newlemmas.append(newi.lower())
        jlemmas = ' '.join(newlemmas)
        for i in Prep:
            if i in jlemmas:
                count += jlemmas.count(i)
        return count / len(words)

    with open('Prep_mw.txt', encoding='utf8') as file:
        lines = file.readlines()
        Prep = [line.rstrip() for line in lines]

    def Prep_mw_pr(words):
        count = 0
        newwords = []
        for i in words:
            newi = str(i)
            newwords.append(newi.lower())
        jwords = ' '.join(newwords)
        for i in Prep:
            if i in jwords:
                count += jwords.count(i)
        return count / len(words)

    with open('Conj_mw.txt', encoding='utf-8') as file:
        lines = file.readlines()
        Conj_mw = [line.rstrip() for line in lines]

    def Conj_mw_pr(words):
        count = 0
        newwords = []
        for i in words:
            newi = str(i)
            newwords.append(newi.lower())
        jwords = ' '.join(newwords)
        for i in Conj_mw:
            if i in jwords:
                count += jwords.count(i)
        return count / len(words)

    with open('LVC.txt', encoding='utf-8') as file:
        lines = file.readlines()
        LVC = [line.rstrip() for line in lines]

    def LVC_pr(lemmas):
        count = 0
        newlemmas = []
        for i in lemmas:
            newi = str(i)
            newlemmas.append(newi.lower())
        jlemmas = ' '.join(newlemmas)
        for i in LVC:
            if i in jlemmas:
                count += jlemmas.count(i)
        return count / len(lemmas)

    with open('Archaic_words.txt', encoding='utf-8') as file:
        lines = file.readlines()
        archaic = [line.rstrip().lower() for line in lines]

    def Arch_pr(lemmas):
        count = 0
        for item in archaic:
            if ' ' in item:
                w = item.split()
                for i in range(len(lemmas) - len(w) + 1):
                    if all(lemmas[i + j] == word for j, word in enumerate(w)):
                        count += 1
            else:
                for word in lemmas:
                    if item == word:
                        count += 1
        return count / len(lemmas)

    N_word = len(words)
    NVR = NVR(upos)
    Autosem_pr = Autosem_pr(upos)
    Func_word_pr = Func_word_pr(upos)
    Verb_pr = Verb_pr(upos)
    Noun_pr = Noun_pr(upos)
    Adj_pr = Adj_pr(upos)
    Pron_pr = Pron_pr(upos)
    Sconj_pr = Sconj_pr(upos)
    Cconj_pr = Cconj_pr(upos)
    Word_form = Word_form(lemmas)
    Gen_pr = Gen_pr(feats)
    Ablt_pr = Ablt_pr(feats)
    dat = dat(feats)
    nomn = nomn(feats)
    loct = loct(feats)
    Neut_pr = Neut_pr(upos, feats)
    Inan_pr = Inan_pr(upos, feats)
    firstp_pr = firstp_pr(upos, feats)
    thirdp_pr = thirdp_pr(upos, feats)
    pres_pr = pres_pr(upos, feats)
    futr_pr = futr_pr(upos, feats)
    past_pr = past_pr(upos, feats)
    impf_pr = impf_pr(upos, feats)
    pf_pr = pf_pr(upos, feats)
    pssv_prtf_pr = pssv_prtf_pr(feats)
    pssv_prts_pr = pssv_prts_pr(feats)
    sya_forms = sya_forms(words, upos)
    yavl_pr = yavl_pr(lemmas)
    abstr_pr = abstr_pr(lemmas)
    deont_pr = deont_pr(lemmas)
    Prep_mw_pr = Prep_mw_pr(words)
    Conj_mw_pr = Conj_mw_pr(words)
    LVC_pr = LVC_pr(lemmas)
    Arch_pr = Arch_pr(lemmas)

    return pd.Series({'N_word': N_word, 'NVR': NVR, 'Autosem_pr': Autosem_pr, 'Func_word_pr': Func_word_pr, 'Verb_pr': Verb_pr,
                      'Noun_pr': Noun_pr, 'Adj_pr': Adj_pr, 'Pron_pr': Pron_pr, 'Sconj_pr': Sconj_pr, 'Cconj_pr': Cconj_pr, 'Word_form': Word_form,
                      'Gen_pr': Gen_pr, 'Ablt_pr': Ablt_pr, 'dat': dat, 'nomn': nomn, 'loct': loct, 'Neut_pr': Neut_pr, 'Inan_pr': Inan_pr,
                      'firstp_pr': firstp_pr, 'thirdp_pr': thirdp_pr, 'pres_pr': pres_pr, 'futr_pr': futr_pr, 'past_pr': past_pr,
                      'impf_pr': impf_pr, 'pf_pr': pf_pr, 'pssv_prtf_pr': pssv_prtf_pr, 'pssv_prts_pr': pssv_prts_pr, 'sya_forms': sya_forms,
                      'yavl_pr': yavl_pr, 'abstr_pr': abstr_pr, 'deont_pr': deont_pr, 'Prep_mw_pr': Prep_mw_pr, 'Conj_mw_pr': Conj_mw_pr,
                      'LVC_pr': LVC_pr, 'Arch_pr': Arch_pr})

metrics = df.groupby('doc_id').apply(calculate_metrics)
metrics.to_excel('metrics_all_docs.xlsx')
