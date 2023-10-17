# Load Required Libraries
library("FactoMineR")
library("xlsx")
library("shiny")
library("Factoshiny")
library("ggplot2")
library("factoextra")
library("corrplot")

# Set Working Directory
setwd("/Chemin  du dossier contenant le fichier xlxs/")

# Load and Preprocess Data
# ------------------------------------

# Load data from an Excel file (adjust the file name and sheet index if needed)
file <- read.xlsx("Dataset.xlsx", sheetIndex = 1, header = TRUE)

# Extract and preprocess the data
df <- file[c(3:22), c(1,5:16)]

# Set column names based on the first row of the data
names(df) <- df[1,]

# Remove the first row and unwanted columns
df <- df[-c(1),]

# Set row names based on the first column of the data
rownames(df) <- df[, 1]

# Remove the first column
df <- df[, -1]

# Reorder columns
df <- subset(df, select = c(12:1))

#Show result of preprocess Dat
View(df)

# Convert data types to numeric
df[] <- lapply(df, type.convert, as.is = TRUE)

# Perform Principal Component Analysis (PCA)
# ------------------------------------

# Perform PCA on the preprocessed data
resultACP <- PCA(df[, 1:12], ncp = 2, scale.unit = TRUE, graph = TRUE)

# Summarize the PCA results
summary(resultACP)
dimdesc(resultACP)

# Access the variance explained by each principal component
resultACP$var

# Access the PCA scores for individuals
resultACP$ind
 
# Calculate eigenvalues of principal components
eig.val <- get_eigenvalue(resultACP)
eig.val
resultACP$eig
res.desc <- dimdesc(resultACP, axes = c(1,2), proba = 0.05)
# Description of dimension 1
res.desc$Dim.1
# Visualize PCA Results
# ------------------------------------

# Scree plot of eigenvalues
fviz_screeplot(resultACP, type = "lines", addlabels = TRUE, main = "Scree Plot PCA")

# Variable plot on the first two dimensions
plot(resultACP, choix = "var", title = "Cercle de corrélation", axes = 1:2)

# Visualize variable contributions and cosines
fviz_pca_var(resultACP, axes = c(1, 2))
fviz_pca_var(resultACP, col.var = "contrib", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)
fviz_pca_var(resultACP, col.var = "cos2", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)

# Correlation plots
custom_colors <- colorRampPalette(c("lightblue", "darkblue"))(100)
corrplot(resultACP$ind$contrib, is.corr = FALSE, col = custom_colors)
corrplot(resultACP$ind$cos2, is.corr = FALSE, cl.ratio = 1)

# Variable and Individual Contributions
# ------------------------------------

# Variable contributions to PC1 and PC2
fviz_contrib(resultACP, choice = "var", axes = 1, top = 12)
fviz_contrib(resultACP, choice = "var", axes = 2, top = 12)
fviz_contrib(resultACP, choice = "var", axes = 1:2, top = 12)

# Individual contributions to PC1 and PC2
fviz_contrib(resultACP, choice = "ind", axes = 1)
fviz_contrib(resultACP, choice = "ind", axes = 2)

# Correlation of variables with PCs and Contributions
corrplot(resultACP$var$cos2, is.corr = FALSE, cl.ratio = 0.5)
corrplot(resultACP$ind$contrib, is.corr = FALSE, cl.ratio = 1)

# K-means Clustering of Variables
# ------------------------------------

# Perform k-means clustering on variable coordinates
set.seed(123)
res.km <- kmeans(resultACP$var$coord, centers = 3, nstart = 25)
grp <- as.factor(res.km$cluster)

# Color variables by cluster groups
fviz_pca_var(resultACP, col.var = grp, palette = c("#0073C2FF", "#EFC000FF", "#868686FF"), legend.title = "Cluster")

# Quality of Variables and Individuals
# ------------------------------------

# Quality of variables on the factor map
fviz_cos2(resultACP, choice = "var", axes = 2, top = 10)

# Quality of individuals on the factor map
fviz_cos2(resultACP, choice = "ind")

# Variable contributions to PC1 and PC2 with color mapping 
fviz_pca_var(resultACP, col.var = "cos2", title = "Cercle de corrélation", gradient.cols = c("yellow", "red"), repel = TRUE)

# Visualize variable contributions to PC1 and PC2
fviz_contrib(resultACP, choice = "var", axes = 1, top = 10)
fviz_contrib(resultACP, choice = "var", axes = 2, top = 10)


# Visualize individual contributions to PC1 and PC2
fviz_contrib(resultACP, choice = "ind", axes = 1:2)

# Individual Plots and Biplot
# ------------------------------------

# Individual plots with color mapping
fviz_pca_ind(resultACP, col.ind = "cos2", title = "Cercle de corr?lation", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)
fviz_pca_ind(resultACP, col.ind = "contrib", title = "Cercle de corr?lation", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)
fviz_pca_ind(resultACP, col.ind = "coord", title = "Cercle de corr?lation", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)

# Biplot of individuals and variables
fviz_pca_biplot(resultACP, geom.ind="point", axes=c(1,2), pointshape=21, pointsize=1,  alpha.var="contrib", col.var= "cos2", gradient.cols=c("#00AFBB", "#E7B800", "#FC4E07"), repel = TRUE)
