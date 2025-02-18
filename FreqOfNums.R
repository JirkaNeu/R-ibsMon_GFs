library(readxl)
library(writexl)

substitute_df = function(){
  print("use random data for illustrating...")
  new_df = data.frame(GFs=NA)
  for (i in 1:25){
    numbers = paste0(c(sample(1:6, sample(1:6, 1), replace=F)), collapse=", ")
    new_df[i, 1] = numbers
  } 
  return(new_df)
}

print_plot = F

gfsdata_df = tryCatch({
  setwd("path_of_file")
  inputfile_jne = "name_of_file.xlsx"
  src_jne = read_xlsx(inputfile_jne, sheet = "Monitoring", col_names = T)
  gfsdata_df = src_jne[21]
  print_plot = T
},
error = function(e) {
  cat("Error: ", conditionMessage(e), "\n")
  print("...")
  gfsdata_df = substitute_df()
},
warning = function(w) {
  cat("Warning: ", conditionMessage(w), "\n")
}
)


gfsdata_df = cbind(gfsdata_df, gf1=NA, gf2=NA, gf3=NA, gf4=NA, gf5=NA, gf6=NA)


for (irow in 1:nrow(gfsdata_df)){
  getnums = as.character(gfsdata_df[irow, 1])
  if(is.na(getnums)){getnums = "ignore this entry"}
  getnums = gsub(" ", "", getnums)
  getnums = strsplit(getnums, ",", fixed=TRUE)[[1]]

  do_sort = T
  for (i in 1:length(getnums)){
    if(getnums[i] != "1" & getnums[i] != "2" & getnums[i] != "3" & getnums[i] != "4" & getnums[i] != "5" & getnums[i] != "6"){
      do_sort = F
      }
  }
  
  if (do_sort == T){
    for (i in 1:length(getnums)){
      switch(getnums[i],
             "1" = {gfsdata_df$gf1[irow] = 1},
             "2" = {gfsdata_df$gf2[irow] = 1},
             "3" = {gfsdata_df$gf3[irow] = 1},
             "4" = {gfsdata_df$gf4[irow] = 1},
             "5" = {gfsdata_df$gf5[irow] = 1},
             "6" = {gfsdata_df$gf6[irow] = 1},
             {NA}#default
      )
    }
  }
}


sum_gfs = NULL
for (i in 1:6){
  sum_gfs = append(sum_gfs, sum(na.omit(gfsdata_df[i+1])))
}

#print(sum_gfs)

light = "#6BA1BB"
dark = "#005478"
Header = "Title of Barchart"
bars = "Column"


#------ base plot ------#
#-----------------------#

p0 = barplot(sum_gfs,
        ylim = c(0, max(sum_gfs)+5),
        col = light,
        main = Header,
        xlab = "",
        ylab = "",
        horiz = F,
        #names.arg = paste(bars, 1:6)
        )
#text(p0, y = sum_gfs+2, labels = c(as.character(sum_gfs)))
text(p0, y = 2, labels = c(paste("GF", 1:6)))
if (print_plot == T){
  dev.print(device = png, filename = 'gfs_plot0.png', width = 1000)
  dev.off()
}


#------ ggplot ------#
#--------------------#

library(ggplot2)
sum_gfs_df = data.frame(GF = c(1:6), freq = sum_gfs, xlabs = c(paste("GF", 1:6)))
long_labels = c("GF 1: First Column",
                "GF 2: Second Column",
                "GF 3: Third Column",
                "GF 4: Fourth Column",
                "GF 5: Fifth Column",
                "GF 6: Sixth Column")
sum_gfs_df = cbind(sum_gfs_df, xlabs_long = long_labels)

p1 = ggplot(sum_gfs_df, aes(x = xlabs, y = freq)) + 
  geom_bar(stat = "identity", color = "black", fill=light) +
  ylim(0, max(sum_gfs)+2) +
  labs(title=Header) +
  geom_text(aes(label=freq), vjust = -1) +
  theme(
    plot.title = element_text(hjust = 0.5, size = 15),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.background = element_blank(),
    #axis.line.x = element_blank(),
    axis.line.x = element_line(colour = "black"),
    axis.line.y = element_blank(),
    axis.title.x=element_blank(), #remove axis title
    axis.title.y=element_blank(), #remove axis title
    #axis.text.x=element_blank(),  #remove axis labels
    axis.text.x = element_text(hjust = 0.5, vjust = 0.5, size = 10),
    axis.text.y=element_blank(),  #remove axis labels
    #axis.ticks.x=element_blank(),  #remove axis ticks
    axis.ticks.y=element_blank()  #remove axis ticks
  )
plot(p1)
if (print_plot == T){
  ggsave("gfs_plot1.png")
}

p2 = ggplot(sum_gfs_df, aes(x = reorder(GF, +freq), y = freq)) + 
  geom_bar(stat = "identity", color = "black", fill=light) +
  labs(title=Header) +
  #geom_text(aes(label=freq), hjust=-1) +
  geom_label(aes(label = long_labels), hjust = 1, nudge_y = -.5, fill = "white", label.size = 0, size = 3) +
  theme(
    plot.title = element_text(hjust = 0.5, size = 15),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.background = element_blank(),
    axis.line.x = element_blank(),
    axis.line.y = element_blank(),
    axis.title.x=element_blank(), #remove axis title
    axis.title.y=element_blank(), #remove axis title
    axis.text.x=element_blank(),  #remove axis labels
    axis.text.y=element_blank(),  #remove axis labels
    axis.ticks.x=element_blank(),  #remove axis ticks
    axis.ticks.y=element_blank()  #remove axis ticks
  ) + 
  coord_flip()
plot(p2)
if (print_plot == T){
  ggsave("gfs_plot2.png")
}

p3 = ggplot(sum_gfs_df, aes(x = reorder(GF, +freq), y = freq)) + 
  geom_bar(stat = "identity", color = "black", fill=light) +
  labs(title=Header) +
  #geom_label(aes(label = long_labels), hjust = 1, nudge_y = -.5, fill = "white", label.size = 0) +
  geom_text(aes(label = long_labels), hjust = 1, nudge_y = -.5, col = "white", size = 3) +
  theme(
    plot.title = element_text(hjust = 0.5, size = 15),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.background = element_blank(),
    axis.line.x = element_blank(),
    axis.line.y = element_blank(),
    axis.title.x=element_blank(), #remove axis title
    axis.title.y=element_blank(), #remove axis title
    axis.text.x=element_blank(),  #remove axis labels
    axis.text.y=element_blank(),  #remove axis labels
    axis.ticks.x=element_blank(),  #remove axis ticks
    axis.ticks.y=element_blank()  #remove axis ticks
  ) + 
  coord_flip()
plot(p3)
if (print_plot == T){
  ggsave("gfs_plot3.png")
}

if (print_plot == T){
  output_jne = "gfs_jne.xlsx"
  write_xlsx(sum_gfs_df, output_jne)
} else {
  cat("Someting went wrong, data not found. \nRandom data generated for illustrating - no output saved.")
}










