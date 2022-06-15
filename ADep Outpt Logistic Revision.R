#This code performs univariate associations between predictor variables and 
#composite of ED visit or hospital admit in 30 days, and multivariable logistic
#regression


#Created by Bradley Fritz on 05/27/2022 (save as from ADep Outpt Logistic)
#Last updated by Bradley Fritz on 05/27/2022
################################################
library(xlsx)
################################################
#Load data
################################################
setwd('\\\\storage1.ris.wustl.edu/bafritz/Active/COVID SSRI')
pts <- read.csv('SSRI Cohort Revision.csv',stringsAsFactors = F)

#Limit to outpatients
pts <- pts[pts$outpt==1,]
################################################
#Prepare the list of exposure variables
################################################
exps <- data.frame(
  name = c('antidepressant','ssri','fluoxetine','fluvoxamine',
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
  label = c('Any Antidepressant','Any SSRI','Fluoxetine','Fluvoxamine',
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
  indent = c(0,0,rep(1,8),0,0,rep(1,6),0,1,1,0,rep(1,4),0,1,1,
             0,rep(1,3),0,rep(1,4)),
  titleonly = c(rep(0,29),1,rep(0,3),1,rep(0,4))
)

exps$indent_string <- ''
exps$indent_string[exps$indent==1] <- '  '
################################################
#Convert dichotomous variables to logical
################################################
vars <- c(exps$name[exps$titleonly==0],'obesity','white','later_time','mood_anx',
          'other_psych','benzo_or_z','antipsychotic')
for (i in 1:length(vars)){
  pts[,vars[i]] <- as.logical(pts[,vars[i]])
}
rm(i,vars)

#Transform continuous variables if needed
pts$abs_ElixIndex <- abs(pts$ElixIndex)

################################################
#Step through exposure variables
################################################
adjglmlist <- list()
outlist <- list()

for (i_exp in 1:dim(exps)[1]){
  print(paste('Performing model ',i_exp,' of ',dim(exps)[1],': ',exps$label[i_exp],sep=''))
  
  #Skip to next value of i_exp if titleonly==1 (sub-heading rather than exposure)
  if (exps$titleonly[i_exp]==1){
    next
  }
  
  #Limit to patients with the exposure or with no antidepressant
  popn <- pts[pts[,which(colnames(pts)==exps$name[i_exp])]==T|pts$antidepressant==F,]
  #Note that in POPN, the exposure variable will have 1:1 correspondence with antidepressant variable
  
  ##############################################
  #Create multivariable logistic regression
  ##############################################
  formulti <- popn[,c('enc_in_30','antidepressant','gender','age',
                      'abs_ElixIndex','nummeds_cat2','obesity','white','mood_anx',
                      'other_psych','benzo_or_z','antipsychotic')]
  
  adjglm <- data.frame(
    summary(glm(enc_in_30~.,family=binomial,data=formulti))$coefficients
  )
  
  tovif <- glm(enc_in_30~.,family=binomial,data=formulti)
  library(DescTools)
  adjvif <- data.frame(VIF(tovif))
  rm(formulti,tovif)
  
  ##############################################
  #Initialize table
  ##############################################
  #Header
  out <- data.frame(Variable = character(),
                    ValidN = character(),
                    Encounter_n = character(),
                    Encounter_pct = character(),
                    Unadj_OR = character(),
                    Unadj_CI = character(),
                    Unadj_p = character(),
                    Adj_OR = character(),
                    Adj_CI = character(),
                    Adj_p = character(),
                    VIF = character())
  #List of rows for categorical vars
  vars <- data.frame(
    name = c('antidepressant','gender','nummeds_cat2',
             'obesity','white','mood_anx','other_psych',
             'benzo_or_z','antipsychotic'),
    label = c(exps$label[i_exp],'Sex',
              'Number of Medications','Obese','White Non-Hispanic',
              'Mood or Anxiety Disorder','Other Psychiatric Disorder',
              'Benzodiazepine or Z Drug','Antipsychotic Drug'))

  ##############################################
  for (i in 1:dim(vars)[1]){
    popn$temp <- as.factor(popn[,vars$name[i]])
    
    freqs <- data.frame(table(popn$temp,popn$enc_in_30))
    coefs <- data.frame(summary(glm(enc_in_30~temp,family=binomial,data=popn))$coefficients)
    ############################################
    #Output summary row
    ############################################
    out <- rbind(out,data.frame(
      Variable = vars$label[i],
      ValidN = '',Encounter_n='',Encounter_pct='',Unadj_OR='',Unadj_CI='',
      #Unadj_p = chisq.test(popn$temp,popn$enc_in_30)$p.value,
      Unadj_p = '',
      Adj_OR='',Adj_CI='',Adj_p='',
      VIF = adjvif$GVIF[rownames(adjvif)==vars$name[i]]
    ))
    
    ############################################
    #Output detail for each level
    ############################################
    lvl <- levels(popn$temp)[1] #Reference level
    out <- rbind(out,data.frame(
      Variable = paste('     ',lvl,sep=''),
      ValidN = sum(freqs$Freq[freqs$Var1==lvl]),
      Encounter_n = sum(freqs$Freq[freqs$Var1==lvl&freqs$Var2==1]),
      Encounter_pct = round(100*sum(freqs$Freq[freqs$Var1==lvl&freqs$Var2==1])/
                                sum(freqs$Freq[freqs$Var1==lvl]),1),
      Unadj_OR = 'ref',Unadj_CI = 'ref',Unadj_p = '',
      Adj_OR = 'ref',Adj_CI='ref',Adj_p = '',VIF = ''
    ))
    rm(lvl)
    
    for (j in 2:length(levels(popn$temp))){
      lvl <- levels(popn$temp)[j]
      
      out <- rbind(out,data.frame(
        Variable = paste('     ',lvl,sep=''),
        ValidN = sum(freqs$Freq[freqs$Var1==lvl]),
        Encounter_n = sum(freqs$Freq[freqs$Var1==lvl&freqs$Var2==1]),
        Encounter_pct = round(100*sum(freqs$Freq[freqs$Var1==lvl&freqs$Var2==1])/
                                  sum(freqs$Freq[freqs$Var1==lvl]),1),
        Unadj_OR = round(exp(coefs$Estimate[rownames(coefs)==paste('temp',lvl,sep='')]),2),
        Unadj_CI = paste(round(exp(coefs$Estimate[rownames(coefs)==paste('temp',lvl,sep='')]-
                                     1.96*coefs$Std..Error[rownames(coefs)==paste('temp',lvl,sep='')]),2),'-',
                         round(exp(coefs$Estimate[rownames(coefs)==paste('temp',lvl,sep='')]+
                                     1.96*coefs$Std..Error[rownames(coefs)==paste('temp',lvl,sep='')]),2),sep=''),
        Unadj_p = coefs$Pr...z..[rownames(coefs)==paste('temp',lvl,sep='')],
        Adj_OR = round(exp(adjglm$Estimate[rownames(adjglm)==paste(vars$name[i],lvl,sep='')]),2),
        Adj_CI = paste(round(exp(adjglm$Estimate[rownames(adjglm)==paste(vars$name[i],lvl,sep='')]-
                                   1.96*adjglm$Std..Error[rownames(adjglm)==paste(vars$name[i],lvl,sep='')]),2),'-',
                       round(exp(adjglm$Estimate[rownames(adjglm)==paste(vars$name[i],lvl,sep='')]+
                                   1.96*adjglm$Std..Error[rownames(adjglm)==paste(vars$name[i],lvl,sep='')]),2),sep=''),
        Adj_p = adjglm$Pr...z..[rownames(adjglm)==paste(vars$name[i],lvl,sep='')],
        VIF=''
      ))
      rm(lvl)
    }
    ############################################
    rm(j,freqs,coefs)
    popn <- popn[,colnames(popn)!='temp']
  }
  rm(i,vars)

  ###################################################
  #List of rows for continuous vars
  vars <- data.frame(
    name = c('age','abs_ElixIndex'),
    label = c('Age','Absolute Value of Elixhauser Index'))
  
  ##############################################
  for (i in 1:dim(vars)[1]){
    popn$temp <- popn[,vars$name[i]]
    
    freqs <- data.frame(table(popn$temp,popn$enc_in_30))
    coefs <- data.frame(summary(glm(enc_in_30~temp,family=binomial,data=popn))$coefficients)
    ############################################
    #Output row
    ############################################
    out <- rbind(out,data.frame(
      Variable = vars$label[i],
      ValidN = sum(freqs$Freq),
      Encounter_n='', Encounter_pct='',
      Unadj_OR=round(exp(coefs$Estimate[rownames(coefs)=='temp']),2),
      Unadj_CI=paste(round(exp(coefs$Estimate[rownames(coefs)=='temp']-
                                 1.96*coefs$Std..Error[rownames(coefs)=='temp']),2),'-',
                     round(exp(coefs$Estimate[rownames(coefs)=='temp']+
                                 1.96*coefs$Std..Error[rownames(coefs)=='temp']),2),sep=''),
      Unadj_p = coefs$Pr...z..[rownames(coefs)=='temp'],
      Adj_OR = round(exp(adjglm$Estimate[rownames(adjglm)==vars$name[i]]),2),
      Adj_CI = paste(round(exp(adjglm$Estimate[rownames(adjglm)==vars$name[i]]-
                                 1.96*adjglm$Std..Error[rownames(adjglm)==vars$name[i]]),2),'-',
                     round(exp(adjglm$Estimate[rownames(adjglm)==vars$name[i]]+
                                 1.96*adjglm$Std..Error[rownames(adjglm)==vars$name[i]]),2),sep=''),
      Adj_p = adjglm$Pr...z..[rownames(adjglm)==vars$name[i]],
      VIF = adjvif$GVIF[rownames(adjvif)==vars$name[i]]
    ))
    ############################################
    rm(freqs,coefs)
    popn <- popn[,colnames(popn)!='temp']
  }
  rm(i,vars)
  
  ##############################################
  #Round p values and VIF
  ##############################################
  for (i in which(colnames(out)%in%c('Unadj_p','Adj_p','VIF'))){
    out$temp <- as.numeric(out[,i])
    
    out[which(out$temp>0.001),i] <- format(round(out$temp[which(out$temp>0.001)],3),nsmall=3)
    out[which(out$temp>0.01),i] <- format(round(out$temp[which(out$temp>0.01)],2),nsmall=2)
    out[which(out$temp<0.001),i] <- paste('< 0.001')
  }
  rm(i)
  
  out <- out[,!colnames(out)%in%'temp']
  ##############################################
  #Save to list
  ##############################################
  adjglmlist[[i_exp]] <- adjglm
  outlist[[i_exp]] <- out
  rm(adjglm,adjvif,out,popn)
  ##############################################
}
rm(i_exp)


################################################
#Create summary table across the exposures
################################################
outsum <- data.frame(
  Exposure = character(),
  N = character(),
  Encounter_n = character(),
  Encounter_pct = character(),
  Unadj_OR = character(),
  Unadj_CI = character(),
  Unadj_p = character(),
  Adj_OR = character(),
  Adj_CI = character(),
  Adj_p = character(),
  VIF = character()
)

#Add the reference group
outsum <- rbind(outsum,data.frame(
  Exposure = 'No Antidepressant',
  N = outlist[[1]]$ValidN[2],
  Encounter_n = outlist[[1]]$Encounter_n[2],
  Encounter_pct = outlist[[1]]$Encounter_pct[2],
  Unadj_OR = 'Ref',Unadj_CI='Ref',Unadj_p = 'Ref',
  Adj_OR = 'Ref',Adj_CI='Ref',Adj_p = 'Ref',VIF = 'NA'
))

for (i_exp in 1:dim(exps)[1]){
  if(exps$titleonly[i_exp]==1){ #S1R and FIASMA headings (no models)
    outsum <- rbind(outsum,data.frame(
      Exposure = paste(exps$indent_string[i_exp],exps$label[i_exp],sep=''),
      N='',Encounter_n='',Encounter_pct='',Unadj_OR='',Unadj_CI='',Unadj_p='',Adj_OR='',Adj_CI='',Adj_p='',VIF=''
    ))
    next
  }
  
  outsum <- rbind(outsum,data.frame(
    Exposure = paste(exps$indent_string[i_exp],exps$label[i_exp],sep=''),
    N = outlist[[i_exp]]$ValidN[3],
    Encounter_n = outlist[[i_exp]]$Encounter_n[3],
    Encounter_pct = outlist[[i_exp]]$Encounter_pct[3],
    Unadj_OR = outlist[[i_exp]]$Unadj_OR[3],
    Unadj_CI = outlist[[i_exp]]$Unadj_CI[3],
    Unadj_p = outlist[[i_exp]]$Unadj_p[3],
    Adj_OR = outlist[[i_exp]]$Adj_OR[3],
    Adj_CI = outlist[[i_exp]]$Adj_CI[3],
    Adj_p = outlist[[i_exp]]$Adj_p[3],
    VIF = outlist[[i_exp]]$VIF[1]
  ))
}
rm(i_exp)
################################################
#Export table
################################################
setwd('\\\\storage1.ris.wustl.edu/bafritz/Active/COVID SSRI')
# write.xlsx(outsum,'COVID AntiDep Outpt Encounter Results Revision.xlsx',row.names = F)
# write.xlsx(outlist[[1]],'COVID AntiDep Outpt Encounter Results Revision.xlsx',
#            sheetName = 'Any Antidepressant',row.names=F,append=T)
