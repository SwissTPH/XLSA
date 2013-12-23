library(ggplot2)
library(reshape2)

setwd("SET_ROOT_DIRECTORY_HERE")

scenarioList <-c("SampleModel")


plotTornado<-function(data,prefix,scenario_id){
  data$range<-abs(data$high-data$low)
  data$sorter<-1/data$range
  data$parname = factor(as.character(data$parname), 
                        levels=levels(data$parname)[order(data$sorter, decreasing=TRUE)])
  tp <- ggplot(data, aes(parname, mode, ymin = low, ymax=high))+ geom_linerange(size=15) + 
				coord_flip()+ylab(levels(data$outcome))+xlab("Parameter")
  win.metafile(file=paste('output/',scenario_id,'/tornado_',prefix,levels(data$outcome),scenario_id,'.wmf',sep=''))
  print(tp)
  dev.off()
}

predictAtQuantiles<-function(data){
  x <- data$parvalue
  y <- data$value
  o <- data$outcome
  n <- data$parname
  mdl<-lm(y ~ x)
  q0<-predict(mdl, data.frame(x = quantile(x,0)))
  q0025<-predict(mdl, data.frame(x = quantile(x,0.025)))
  q05<-predict(mdl, data.frame(x = quantile(x,0.5)))
  q0975<-predict(mdl, data.frame(x = quantile(x,0.975)))
  q1<-predict(mdl, data.frame(x = quantile(x,1)))
  c(parname=levels(droplevels(data$parname)),outcome=levels(droplevels(data$outcome)),high=q0975,low=q0025,mode=q05,qzero=q0,qone=q1)
}

doPSA<-function(data,scenario_id){
  PSAByParam<-split(data,data$parname)
  df<-NULL
  for (p in PSAByParam){
    df<-rbind(df,(predictAtQuantiles(p)))
  }
  df<-data.frame(df, scenario_id)
  df$low<-as.numeric(as.character(df$low.2.5.))
  df$high<-as.numeric(as.character(df$high.97.5.))
  df$mode<-as.numeric(as.character(df$mode.50.))
  write.table(df,paste('output/',scenario_id,"/",levels(df$outcome),scenario_id,".txt", sep=""), sep = "\t", row.names = FALSE)
  plotTornado(df,'psa',scenario_id)
}

for (scenario_id in scenarioList){
  dir.create(file.path(paste('output/',scenario_id, sep='')), showWarnings = FALSE)
  PSA <- read.table(paste('output/PSA_',scenario_id,'_long.txt',sep=''), header=TRUE, sep="\t", na.strings="NA", dec=".", strip.white=TRUE)
  oneway <- read.table(paste('output/univariate_',scenario_id,'_long.txt',sep=''), header=TRUE, sep="\t", na.strings="NA", dec=".", strip.white=TRUE)
  oneway_wide<-dcast(oneway,parname+ outcome ~ assumption, value="value")
  byOutcome<-split(PSA,PSA$outcome)
  sapply(byOutcome,doPSA,scenario_id)
}