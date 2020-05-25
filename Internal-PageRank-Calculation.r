library("igraph")
# Swap out path to your Screaming Frog All Outlink CSV. For Windows, remember to change backslashes to forward slashes.
links <- read.csv("C:/Users/evgeniy/Documents/si_all_outlinks.csv", skip = 1) # CSV Path
# This line of code is optional. It filters out JavaScript, CSS, and Images. Technically you should keep them in there.
links <- subset(links, Type=="AHREF") # Optional line. Filter.
links <- subset(links, Follow=="true")
links <- subset(links, select=c(Source,Destination))
g <- graph.data.frame(links)
pr <- page.rank(g, algo = "prpack", vids = V(g), directed = TRUE, damping = 0.85)
values <- data.frame(pr$vector)
values$names <- rownames(values)
row.names(values) <- NULL
values <- values[c(2,1)]
names(values)[1] <- "url"
names(values)[2] <- "pr"
# Swap out 'domain' and 'com' to represent your website address.
values <- values[grepl("https?:\\/\\/(.*\\.)?signal-iduna\\.de.*", values$url),] # Domain filter.
# Replace with your desired filename for the output file.
write.csv(values, file = "si-output-pagerank.csv") # Output file.