#This code performs descriptive statistics for patients taking antidepressant
#versus no antidepressant

#06/07/22 - Add non-antidepressant FIASMA and S1R to table


#Created by Bradley Fritz on 05/27/2022 (save as from ADep Outpt Description.R)
#Last updated by Bradley Fritz on 06/07/2022
################################################
library(xlsx)
################################################
#Load data
################################################
setwd('\\\\storage1.ris.wustl.edu/bafritz/Active/COVID SSRI')
pts <- read.csv('SSRI Cohort Revision.csv',stringsAsFactors = F)

#Limit to outpatients
pts <- pts[pts$outpt==1,]
#Recode antidepressant variable
pts$antidepressant <- as.logical(pts$antidepressant)
################################################
#Header for table and N
################################################
#Initialize data frame
desc <- data.frame(Variable = character(),
                   nMiss = character(),
                   Overall.n = integer(),
                   Overall.pct = character(),
                   AntiDep.n = integer(),
                   AntiDep.pct = character(),
                   NoAntiDep.n = integer(),
                   NoAntiDep.pct = character(),
                   p = character())

#N
N1 <- dim(pts[pts$antidepressant==1,])[1]
N0 <- dim(pts[pts$antidepressant==0,])[1]
desc <- rbind(desc,data.frame(
  Variable = 'N',nMiss = '',
  Overall.n = N1+N0,
  Overall.pct = '',
  AntiDep.n = N1,
  AntiDep.pct = round(100*N1/(N1+N0),1),
  NoAntiDep.n = N0,
  NoAntiDep.pct = round(100*N0/(N1+N0),1),
  p = ''
))
###############################################
#Description for continuous variables
###############################################
vars <- data.frame(
  name = c('age','ElixIndex'),
  label = c('Age (yr)','Elixhauser Index')
)

for (i in 1:dim(vars)[1]){
  
  desc <- rbind(desc,data.frame(
    Variable = vars$label[i],
    nMiss = sum(is.na(pts[,vars$name[i]])),
    Overall.n = fivenum(pts[,vars$name[i]])[3],
    Overall.pct = paste(fivenum(pts[,vars$name[i]])[2],'to',
                        fivenum(pts[,vars$name[i]])[4]),
    AntiDep.n = fivenum(pts[pts$antidepressant==1,vars$name[i]])[3],
    AntiDep.pct = paste(fivenum(pts[pts$antidepressant==1,vars$name[i]])[2],'to',
                        fivenum(pts[pts$antidepressant==1,vars$name[i]])[4]),
    NoAntiDep.n = fivenum(pts[pts$antidepressant==0,vars$name[i]])[3],
    NoAntiDep.pct = paste(fivenum(pts[pts$antidepressant==0,vars$name[i]])[2],'to',
                          fivenum(pts[pts$antidepressant==0,vars$name[i]])[4]),
    p = wilcox.test(pts[pts$antidepressant==0,vars$name[i]],
                    pts[pts$antidepressant==1,vars$name[i]])$p.value
  ))
}
rm(i,vars)

###############################################
#Description for multi-level categorical variables
###############################################
vars <- data.frame(
  name = c('gender','race','ethnicity','nummeds_cat2'),
  label = c('Sex','Race','Ethnicity','Number of Home Medications'))

for (i in 1:dim(vars)[1]){
  pts[,vars$name[i]] <- as.factor(pts[,vars$name[i]])
  
  freq <- data.frame(table(pts[,vars$name[i]]))
  cross <- data.frame(table(pts$antidepressant,pts[,vars$name[i]]))
  
  #Output summary row
  desc <- rbind(desc,data.frame(
    Variable = vars$label[i],
    nMiss = (N0+N1)-sum(freq$Freq),
    Overall.n='',Overall.pct='',AntiDep.n = '',
    AntiDep.pct='',NoAntiDep.n = '',NoAntiDep.pct='',
    p = chisq.test(pts[,vars$name[i]],pts$antidepressant)$p.value
  ))
  
  #Output detail for each level
  for (j in 1:dim(freq)[1]){
    desc <- rbind(desc,data.frame(
      Variable = paste('     ',freq$Var1[j],sep=''),
      nMiss = '',
      Overall.n = freq$Freq[j],
      Overall.pct = round(100*freq$Freq[j]/sum(freq$Freq),1),
      AntiDep.n = cross$Freq[cross$Var1==T&cross$Var2==freq$Var1[j]],
      AntiDep.pct = round(100*cross$Freq[cross$Var1==T&cross$Var2==freq$Var1[j]]/
                              sum(cross$Freq[cross$Var1==T]),1),
      NoAntiDep.n = cross$Freq[cross$Var1==F&cross$Var2==freq$Var1[j]],
      NoAntiDep.pct = round(100*cross$Freq[cross$Var1==F&cross$Var2==freq$Var1[j]]/
                                sum(cross$Freq[cross$Var1==F]),1),
      p=''
    ))
  }
  rm(j,freq,cross)
}
rm(i,vars)

################################################
#Description for dichotomous variables
################################################
vars <- data.frame(name=c('obesity','white','mood_anx',
                          'other_psych','benzo_or_z','antipsychotic',
                          'nonadfiasma','nonads1r'),
                   label=c('Obese','White Non-Hispanic',
                           'Mood or Anxiety Disorder',
                           'Other Psychiatric Disorder',
                           'Any Benzodiazepine or Z Drug','Any Antipsychotic',
                           'Non-Antidepressant FIASMA Drug',
                           'Non-Antidepressant S1R Drug'))

for (i in 1:dim(vars)[1]){
  pts[,vars$name[i]] <- as.integer(pts[,vars$name[i]])
  
  freq <- data.frame(table(pts[,vars$name[i]]))
  cross <- data.frame(table(pts$antidepressant,pts[,vars$name[i]]))
  
  #Output row
  desc <- rbind(desc,data.frame(
    Variable = vars$label[i],
    nMiss = (N0+N1)-sum(freq$Freq),
    Overall.n = freq$Freq[freq$Var1==1],
    Overall.pct = round(100*freq$Freq[freq$Var1==1]/sum(freq$Freq),1),
    AntiDep.n = cross$Freq[cross$Var1==T&cross$Var2==1],
    AntiDep.pct = round(100*cross$Freq[cross$Var1==T&cross$Var2==1]/
                            sum(cross$Freq[cross$Var1==T]),1),
    NoAntiDep.n = cross$Freq[cross$Var1==F&cross$Var2==1],
    NoAntiDep.pct = round(100*cross$Freq[cross$Var1==F&cross$Var2==1]/
                              sum(cross$Freq[cross$Var1==F]),1),
    p = chisq.test(pts[,vars$name[i]],pts$antidepressant)$p.value
  ))
  rm(freq,cross)
}
rm(i,vars)

# ################################################
# #Medications
# ################################################
# 
# vars <- data.frame(name=c('benzo_or_z','alprazolam','chlordiazepoxide','clorazepate',
#                           'diazepam','flurazepam','lorazepam','oxazepam',
#                           'temazepam','triazolam','eszopiclone','zaleplon','zolpidem',
#                           'antipsychotic','aripiprazole','asenapine','brexpiprazole',
#                           'cariprazine','chlorpromazine','clozapine','droperidol',
#                           'fluphenazine','haloperidol','iloperidone','loxapine',
#                           'lurasidone','olanzapine','paliperidone','perphenazine',
#                           'pimavanserin','pimozide','quetiapine','risperidone',
#                           'thiothiene','ziprasidone',
#                           'ssri','ssri_high','fluoxetine','fluvoxamine',
#                           'ssri_int','citalopram','escitalopram',
#                           'ssri_low','paroxetine','ssri_antag','sertraline',
#                           'vilazodone','vortioxetine',
#                           'tca','amitriptyline','clomipramine','desipramine',
#                           'doxepin','imipramine','nortriptyline',
#                           'phenylpiperazine','nefazodone','trazodone',
#                           'snri','desvenlafaxine','duloxetine','levomilnacipran',
#                           'venlafaxine','bupropion','mirtazapine'),
#                    label=c('Any Benzodiazepine or Z Drug',
#                            'Alprazolam','Chlordiazepoxide','Clorazepate',
#                            'Diazepam','Flurazepam','Lorazepam','Oxazepam',
#                            'Temazepam','Triazolam','Eszopiclone','Zaleplon','Zolpidem',
#                            'Any Antipyschotic','Arirpiprazole','Asenapine',
#                            'Brexpiprazole','Cariprazine','Chlorpromazine','Clozapine',
#                            'Droperidol','Fluphenazine','Haloperidol','Iloperidone',
#                            'Loxapine','Lurasidone','Olanzapine','Paliperidone',
#                            'Perphenazine','Pimavanserin','Pimozide','Quetiapine',
#                            'Risperidone','Thiothiene','Ziprasidone',
#                            'Any SSRI','SSRI with High S1R Affinity (Agonist)',
#                            'Fluoxetine','Fluvoxamine',
#                            'SSRI with Intermediate S1R Affinity',
#                            'Citalopram','Escitalopram',
#                            'SSRI with Low S1R Affinity','Paroxetine',
#                            'SSRI with High S1R Affinity (Antagonist)','Sertraline',
#                            'Vilazodone','Vortioxetine',
#                            'Any TCA','Amitriptyline','Clomipramine',
#                            'Desipramine','Doxepin','Imipramine','Nortriptyline',
#                            'Any Phenylpiperazine','Nefazodone','Trazodone',
#                            'Any SNRI','Desvenlafaxine','Duloxetine',
#                            'Levomilnacipran','Venlafaxine',
#                            'Bupropion','Mirtazapine'),
#                    category=c(1,rep(0,12),1,rep(0,21),rep(1,2),rep(0,2),1,rep(0,2),1,0,1,0,rep(1,3),
#                               rep(0,6),1,rep(0,2),1,rep(0,4),rep(1,2)),
#                    antidep=c(rep(0,35),rep(1,30)))
# 
# for (i in 1:dim(vars)[1]){
#   if (vars$category[i]==0){
#     vars$label[i] <- paste('     ',vars$label[i],sep='')
#   }
#   
#   pts[,vars$name[i]] <- as.integer(pts[,vars$name[i]])
#   
#   freq <- data.frame(table(pts[,vars$name[i]]))
#   cross <- data.frame(table(pts$antidepressant,pts[,vars$name[i]]))
#   
#   if (dim(freq)[1]==1)next
#   
#   if (vars$antidep[i]==1){
#     # #Output row if this is an antidepressant. (Omit the p value)
#     # desc <- rbind(desc,data.frame(
#     #   Variable = vars$label[i],
#     #   nMiss = (N0+N1)-sum(freq$Freq),
#     #   Overall = paste(freq$Freq[freq$Var1==1],' (',
#     #                   round(100*freq$Freq[freq$Var1==1]/sum(freq$Freq)),'%)',sep=''),
#     #   AntiDep = paste(cross$Freq[cross$Var1==T&cross$Var2==1],' (',
#     #                   round(100*cross$Freq[cross$Var1==T&cross$Var2==1]/
#     #                           sum(cross$Freq[cross$Var1==T])),'%)',sep=''),
#     #   NoAntiDep = '',p = ''
#     # ))
#   }else {
#     #Output row
#     desc <- rbind(desc,data.frame(
#       Variable = vars$label[i],
#       nMiss = (N0+N1)-sum(freq$Freq),
#       Overall = paste(freq$Freq[freq$Var1==1],' (',
#                       round(100*freq$Freq[freq$Var1==1]/sum(freq$Freq)),'%)',sep=''),
#       AntiDep = paste(cross$Freq[cross$Var1==T&cross$Var2==1],' (',
#                       round(100*cross$Freq[cross$Var1==T&cross$Var2==1]/
#                               sum(cross$Freq[cross$Var1==T])),'%)',sep=''),
#       NoAntiDep = paste(cross$Freq[cross$Var1==F&cross$Var2==1],' (',
#                         round(100*cross$Freq[cross$Var1==F&cross$Var2==1]/
#                                 sum(cross$Freq[cross$Var1==F])),'%)',sep=''),
#       p = chisq.test(pts[,vars$name[i]],pts$antidepressant)$p.value
#     ))
#   }
# 
#   rm(freq,cross)
# }
# rm(i,vars)

################################################
#Round p values
################################################
for (i in which(colnames(desc)%in%'p')){
  desc$temp <- as.numeric(desc[,i])
  
  desc[which(desc$temp>0.001),i] <- format(round(desc$temp[which(desc$temp>0.001)],3),nsmall=3)
  desc[which(desc$temp>0.01),i] <- format(round(desc$temp[which(desc$temp>0.01)],2),nsmall=2)
  desc[which(desc$temp<0.001),i] <- paste('< 0.001')
}
rm(i)

desc <- desc[,!colnames(desc)%in%'temp']

################################################
#Table describing the antidepressants used
################################################
#Initialize table
adep <- data.frame(
  Antidepressant = character(),
  N = character(),
  Dose_Median = character(),
  Dose_IQR = character(),
  Dose_Range = character(),
  Fluox_Eq_Median = character(),
  Fluox_Eq_IQR = character(),
  Fluox_Eq_Range = character()
)


vars <- data.frame(name=c('antidepressant','ssri','fluoxetine','fluvoxamine',
                          'citalopram','escitalopram','paroxetine','sertraline',
                          'vilazodone','vortioxetine',
                          'nonssri','tca','amitriptyline','clomipramine',
                          'desipramine','doxepin','imipramine','nortriptyline',
                          'phenylpiperazine','nefazodone','trazodone',
                          'snri','desvenlafaxine','duloxetine','levomilnacipran',
                          'venlafaxine',
                          'other_adep','bupropion','mirtazapine',
                          'NULL','ssri_high','ssri_int','ssri_low',
                          'NULL','fiasma','unk_fiasma','unk_fiasma_nobup',
                          'unk_fiasma_nobupmirt'),
                   label=c('Any Antidepressant','Any SSRI','Fluoxetine','Fluvoxamine',
                           'Citalopram','Escitalopram','Paroxetine','Sertraline',
                           'Vilazodone','Vortioxetine',
                           'Any Non-SSRI','Any TCA','Amitriptyline','Clomipramine',
                           'Desipramine','Doxepin','Imipramine','Nortriptyline',
                           'Any Phenylpiperazine','Nefazodone','Trazodone',
                           'Any SNRI','Desvenlafaxine','Duloxetine',
                           'Levomilnacipran','Venlafaxine',
                           'Other Antidepressant','Bupropion','Mirtazapine',
                           'SSRIs GROUPED BY S1R AFFINITY',
                           'High Affinity','Intermediate Affinity','Low Affinity',
                           'GROUPED BY FIASMA',
                           'Antidepressants with FIASMA',
                           'Antidepressants with Unknown FIASMA',
                           'Antidepressants with Unknown FIASMA (excluding bupropion)',
                           'Antidepressants with Unknown FIASMA (excluding bupropion and mirtazapine)'),
                   hasdose=c(0,0,rep(1,8),0,0,rep(1,6),0,1,1,0,rep(1,4),0,1,1,
                             rep(0,9)),
                   indent=c(0,0,rep(1,8),0,0,rep(1,6),0,1,1,0,rep(1,4),0,1,1,
                            0,rep(1,3),0,rep(1,4)),
                   titleonly=c(rep(0,29),1,rep(0,3),1,rep(0,4)))

vars$indent_string <- ''
vars$indent_string[vars$indent==1] <- '  '

for (i in 1:dim(vars)[1]){
  if(vars$titleonly[i]==1){ #Titles of the S1R and FIASMA Sections
    adep <- rbind(adep,data.frame(
      Antidepressant = paste(vars$indent_string[i],vars$label[i],sep=''),
      N = '',Dose_Median='',Dose_IQR='',Dose_Range='',Fluox_Eq_Median='',Fluox_Eq_IQR='',Fluox_Eq_Range=''
    ))
  } else{
    
    col_dich <- which(colnames(pts)==vars$name[i])
    col_fluox <- which(colnames(pts)==paste(vars$name[i],'_dose_fluox_eq',sep=''))
    
    if(vars$hasdose[i]==0){ #Drug groups (no individual dosages, but can have fluox eqvts)
      adep <- rbind(adep,data.frame(
        Antidepressant = paste(vars$indent_string[i],vars$label[i],sep=''),
        N = sum(pts[,col_dich]),
        Dose_Median = '',Dose_IQR='',Dose_Range = '',
        Fluox_Eq_Median = fivenum(pts[,col_fluox])[3],
        Fluox_Eq_IQR = paste(fivenum(pts[,col_fluox])[2],'-',
                             fivenum(pts[,col_fluox])[4],sep=''),
        Fluox_Eq_Range = paste(fivenum(pts[,col_fluox])[1],' to ',
                               fivenum(pts[,col_fluox])[5],sep='')
      ))
    } else{ #Individual drug, can report dosages
      
      col_dose <- which(colnames(pts)==paste(vars$name[i],'_dose',sep=''))
      
      adep <- rbind(adep,data.frame(
        Antidepressant = paste(vars$indent_string[i],vars$label[i],sep=''),
        N = sum(pts[,col_dich]),
        Dose_Median = fivenum(pts[,col_dose])[3],
        Dose_IQR = paste(fivenum(pts[,col_dose])[2],'-',
                         fivenum(pts[,col_dose])[4],sep=''),
        Dose_Range = paste(fivenum(pts[,col_dose])[1],' to ',
                           fivenum(pts[,col_dose])[5],sep=''),
        Fluox_Eq_Median = fivenum(pts[,col_fluox])[3],
        Fluox_Eq_IQR = paste(fivenum(pts[,col_fluox])[2],'-',
                             fivenum(pts[,col_fluox])[4],sep=''),
        Fluox_Eq_Range = paste(fivenum(pts[,col_fluox])[1],' to ',
                               fivenum(pts[,col_fluox])[5],sep='')
      ))
      rm(col_dose)
    }
    rm(col_dich,col_fluox)
    
  }
}
rm(i)



################################################
#Export table
################################################
setwd('\\\\storage1.ris.wustl.edu/bafritz/Active/COVID SSRI')
# write.xlsx(desc,'COVID AntiDep Outpt Descriptive Statistics Revision.xlsx',sheetName = 'Two Groups',row.names = F)
# write.xlsx(adep,'COVID AntiDep Outpt Descriptive Statistics Revision.xlsx',sheetName = 'Antidepressants',row.names = F,append=T)