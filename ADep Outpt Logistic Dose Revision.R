#This code performs univariate associations between predictor variables and 
#composite of ED visit or hospital admit in 30 days, and multivariable logistic
#regression. Compares low, medium, and high fluoxetine equivalent dosing

#06/07/22 - Add non-AD FIASMA/S1R to model

#Created by Bradley Fritz on 06/03/2022 via save as from ADep Outpt Logistic Dose.R
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
pts$abs_ElixIndex <- abs(pts$ElixIndex)

################################################
#Prepare the list of exposure variables
################################################
exps <- data.frame(
  name = c('antidepressant','ssri','bupropion','fiasma'),
  label = c('Any Antidepressant','Any SSRI','Bupropion',
            'Antidepressants with FIASMA'),
  indent = rep(0,4),
  maintext = rep(1,4) #Distinguish Table 4 (main text) from supplement table
)

################################################
#Convert dichotomous variables to logical
################################################
vars <- c(exps$name,'antidepressant','obesity','white','mood_anx',
          'other_psych','benzo_or_z','antipsychotic')
for (i in 1:length(vars)){
  pts[,vars[i]] <- as.logical(pts[,vars[i]])
}
rm(i,vars)

#Calculate composite of non-AD FIASMA and S1R
pts$nonad_fs1r <- pts$nonadfiasma | pts$nonads1r

################################################
#Step through exposure variables
################################################
adjglmlist <- list()
outlist <- list()

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

for (i_exp in 1:dim(exps)[1]){
  print(paste('Performing model ',i_exp,' of ',dim(exps)[1],': ',exps$label[i_exp],sep=''))
  
  #Header for this exposure
  outsum <- rbind(outsum,data.frame(
    Exposure = exps$label[i_exp],
    N='',Encounter_n='',Encounter_pct='',Unadj_OR='',Unadj_CI='',Unadj_p='',
    Adj_OR='',Adj_CI='',Adj_p='',VIF=''
  ))
  
  #Reference Group
  temp <- data.frame(table(pts$antidepressant,pts$enc_in_30))
  
  outsum <- rbind(outsum,data.frame(
    Exposure = 'No Antidepressant',
    N = sum(temp$Freq[temp$Var1==F]),
    Encounter_n = temp$Freq[temp$Var1==F&temp$Var2==1],
    Encounter_pct = round(100*temp$Freq[temp$Var1==F&temp$Var2==1]/sum(temp$Freq[temp$Var1==F]),1),
    Unadj_OR='Ref',Unadj_CI='Ref',Unadj_p='Ref',
    Adj_OR='Ref',Adj_CI='Ref',Adj_p='Ref',VIF='Ref'
  ))
  rm(temp)
  
  #Limit to patients with the exposure or with no antidepressant
  popn <- pts[pts[,which(colnames(pts)==exps$name[i_exp])]==T|pts$antidepressant==F,]
  #Note that in POPN, the exposure variable will have 1:1 correspondence with antidepressant variable
  
  ##############################################
  #Create categorical variable for fluoxetine equivalents
  ##############################################
  dose_col <- which(colnames(popn)==paste(exps$name[i_exp],'_dose_fluox_eq',sep='')) #Identify column
  
  popn$f_eqvts_lt20 <- popn[,dose_col] < 20
  popn$f_eqvts_gte20 <- popn[,dose_col] >= 20
  popn$f_eqvts_20to39 <- (popn[,dose_col] >= 20) & (popn[,dose_col] < 40)
  popn$f_eqvts_gte40 <- popn[,dose_col] >= 40
  rm(dose_col)
  
  ##############################################
  #Step through the four fluoxetine eqvts exposures
  ##############################################
  feqs <- data.frame(name=c('f_eqvts_lt20','f_eqvts_gte20','f_eqvts_20to39','f_eqvts_gte40'),
                     label=c('< 20 mg Fluoxetine Equivalent',
                             '>= 20 mg Fluoxetine Equivalent',
                             '20-39 mg Fluoxetine Equivalent',
                             '>= 40 mg Fluoxetine Equivalent'))
  
  for (k in 1:dim(feqs)[1]){
    ##############################################
    #Create multivariable logistic regression
    ##############################################
    
    formulti <- popn[popn[,which(colnames(popn)==feqs$name[k])]|
                       (popn$antidepressant==F),
                     c('enc_in_30','antidepressant','gender','age',
                        'abs_ElixIndex','nummeds_cat2','obesity','white','mood_anx',
                        'other_psych','benzo_or_z','antipsychotic','nonad_fs1r')]
    
    adjglm <- data.frame(
      summary(glm(enc_in_30~.,family=binomial,data=formulti))$coefficients
    )
    
    tovif <- glm(enc_in_30~.,family=binomial,data=formulti)
    library(DescTools)
    adjvif <- data.frame(VIF(tovif))
    rm(tovif)
    
    ##############################################
    #Initialize table
    ##############################################
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
    
    ###############################################
    #List of rows for categorical vars
    vars <- data.frame(
      name = c('antidepressant','gender','nummeds_cat2',
               'obesity','white','mood_anx','other_psych',
               'benzo_or_z','antipsychotic','nonad_fs1r'),
      label = c(feqs$label[k],'Sex',
                'Number of Medications','Obese','White Non-Hispanic',
                'Mood or Anxiety Disorder','Other Psychiatric Disorder',
                'Benzodiazepine or Z Drug','Antipsychotic Drug',
                'Non-Antidepressant FIASMA or S1R'))

    ##############################################
    for (i in 1:dim(vars)[1]){
      formulti$temp <- as.factor(formulti[,vars$name[i]])
      
      freqs <- data.frame(table(formulti$temp,formulti$enc_in_30))
      coefs <- data.frame(summary(glm(enc_in_30~temp,family=binomial,data=formulti))$coefficients)
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
      lvl <- levels(formulti$temp)[1] #Reference level
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
      
      for (j in 2:length(levels(formulti$temp))){
        lvl <- levels(formulti$temp)[j]
        
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
      formulti <- formulti[,colnames(formulti)!='temp']
    }
    rm(i,vars)
    
    ###################################################
    #List of rows for continuous vars
    vars <- data.frame(
      name = c('age','abs_ElixIndex'),
      label = c('Age','Absolute Value of Elixhauser Index'))
    
    ##############################################
    for (i in 1:dim(vars)[1]){
      formulti$temp <- formulti[,vars$name[i]]
      
      freqs <- data.frame(table(formulti$temp,formulti$enc_in_30))
      coefs <- data.frame(summary(glm(enc_in_30~temp,family=binomial,data=formulti))$coefficients)
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
    rm(i,vars,formulti)
    
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
    #Add to summary table
    ##############################################
    outsum <- rbind(outsum,data.frame(
      Exposure = out$Variable[1],
      N = out$ValidN[3],
      Encounter_n = out$Encounter_n[3],
      Encounter_pct = out$Encounter_pct[3],
      Unadj_OR = out$Unadj_OR[3],
      Unadj_CI = out$Unadj_CI[3],
      Unadj_p = out$Unadj_p[3],
      Adj_OR = out$Adj_OR[3],
      Adj_CI = out$Adj_CI[3],
      Adj_p = out$Adj_p[3],
      VIF = out$VIF[1]
      ))
    
    ##############################################
    #Save to list
    ##############################################
    adjglmlist[[i_exp*10+k]] <- adjglm
    outlist[[i_exp*10+k]] <- out
    rm(adjglm,adjvif,out)
  
  }
  rm(k,feqs,popn)
  
  

}
rm(i_exp)


################################################

################################################
#Export table
################################################
setwd('\\\\storage1.ris.wustl.edu/bafritz/Active/COVID SSRI')
# write.xlsx(outsum,'COVID AntiDep Outpt Encounter Dose Results Revision.xlsx',row.names = F)