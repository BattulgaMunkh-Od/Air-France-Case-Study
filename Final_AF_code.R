# Load necessary libraries
library(readxl)             #For loading excel data set
library(readr)              #For loading excel data set
library(dplyr)              #For data manipulation
library(ggplot2)            #For enhanced visualization
library(fastDummies)        #For creating dummy variables on categorical variables
library(splitstackshape)    #For structured split and train
library(caret)              #For Confusion matrix
library(rpart)              #For generating gini tree
library(rpart.plot)         #For generating gini tree
library(ROCR)               #For AUC ROC analysis

#importing our data set 
file <- read_excel("D:/tulgaa/Hult/MBAN/Spring/R/Air France Case Spreadsheet Supplement.xls",sheet="DoubleClick")

my_data <- as.data.frame(file)

# describing the data frame.
print(head(my_data))


# We utilized dplyr we learned in Pavel's class!

# Summary grouped by Publisher Name
publisher_analysis <- my_data %>%
  group_by(`Publisher Name`) %>% # Group data by publisher name
  summarise(
    Total_Clicks = sum(Clicks), # Sum of clicks
    Total_Click_Charges = sum(`Click Charges`), # Sum of clicks
    Total_Impressions = sum(Impressions), # Sum of Impressions
    
    Total_Net_Revenue = sum(Amount) - Total_Click_Charges,  # Total net revenue after deducting click charges
    Avg_Cost_Per_Click = Total_Click_Charges / Total_Clicks, # Average cost per click (CPC)
    Total_Bookings = sum(`Total Volume of Bookings`), # Total number of bookings
    Revenue_Per_Booking = sum(Amount) / Total_Bookings, # Revenue generated per booking
    ROA = Total_Net_Revenue / Total_Click_Charges * 100, # Return on Ad Spend (ROA), expressed as a percentage
    Conversion_Rate_Impressions = Total_Bookings / Total_Impressions * 100, # Conversion rate: percentage of impressions that resulted in bookings
    Conversion_Rate_Clicks = Total_Bookings / Total_Clicks * 100, # Conversion rate: percentage of clicks that resulted in bookings
    Total_Click_Costs = Total_Click_Charges, # Total cost incurred for clicks
    Cost_Per_Booking = Total_Click_Costs / Total_Bookings # Cost incurred per booking
  ) %>%
  arrange(`Publisher Name`) # Sort results by Publisher Name in ascending order

print(as.data.frame(publisher_analysis), row.names = FALSE)


# Plot the distribution of Conversion
my_data$Conversion_Booking <- my_data$`Total Volume of Bookings`


my_data$Conversion_Booking[my_data$Conversion_Booking > 0] <- 1
my_data$Conversion_Booking[my_data$Conversion_Booking == 0] <- 0

#generating visuals for the Distribution conversion
ggplot(my_data, aes(x = Conversion_Booking)) +
  geom_histogram(bins = 30, fill = "blue", alpha = 0.7) +
  facet_wrap(~`Publisher Name`) +
  labs(title = "Distribution of Conversion",
       x = "Conversion = 1",
       y = "Count") +
  theme_minimal()

# Summary grouped by Campaign and Keyword
my_data %>%
  group_by(Campaign, Keyword) %>% # Group data by publisher name and keyword
  summarise(
    Total_Net_Revenue = sum(Amount)  # Total net revenue after deducting click charges
  )


# Check the structure of the dataset
str(my_data)

# Summary statistics of the dataset
summary(my_data)

# Display the first few rows
head(my_data)

##Focus on Revenue##

# Total revenue by Publisher Name
my_data %>%
  group_by(`Publisher Name`) %>%
  summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
  arrange(desc(Total_Revenue))

my_data %>%
  group_by(`Match Type`) %>%
  summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
  arrange(desc(Total_Revenue))

# Total revenue by campaign
my_data %>%
  group_by(Campaign) %>%
  summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
  arrange(desc(Total_Revenue))

##Reduce Costs (CPC, Keyword Analysis)##

# Plot the distribution of CPC
ggplot(my_data, aes(x = `Avg. Cost per Click`)) +
  geom_histogram(bins = 30, fill = "blue", alpha = 0.7) +
  labs(title = "Distribution of Cost Per Click", x = "Cost Per Click", y = "Count")

# Top 10 keywords with the highest avg CPC & avg conversion rate
my_data %>%
  group_by(Keyword) %>%  # Group by Keyword
  summarise(
    Avg_CPC = mean(`Avg. Cost per Click`, na.rm = TRUE),  # Calculate mean CPC
    Total_Bookings = sum(`Total Volume of Bookings`, na.rm = TRUE),  # Sum of bookings
    Total_Clicks = sum(Clicks, na.rm = TRUE),  # Sum of clicks
    Conversion_Rate = ifelse(Total_Clicks > 0, Total_Bookings / Total_Clicks, NA)  # Calculate conversion rate
  ) %>%
  filter(Total_Bookings > 0) %>%  # Exclude rows where Total_Bookings is 0
  arrange(desc(Avg_CPC)) %>%  # Sort by highest Avg_CPC
  print(n=100)  # Get top 100 keywords


##Drive Traffic to Website (Ad Campaign Effectiveness)##

# Total clicks by campaign
my_data %>%
  group_by(Campaign) %>%
  summarise(Total_Clicks = sum(Clicks, na.rm = TRUE)) %>%
  arrange(desc(Total_Clicks))

# Average Click-Through Rate (CTR) by publisher
my_data %>%
  group_by(`Publisher Name`) %>%
  summarise(
    Avg_Click_Through_Rate = mean(`Engine Click Thru %`, na.rm = TRUE),  # Calculate average CTR
    Total_Bookings = sum(`Total Volume of Bookings`, na.rm = TRUE),  # Sum of bookings
    Total_Clicks = sum(Clicks, na.rm = TRUE),  # Sum of clicks (Check if column name is "Click" or "Clicks")
    Conversion_Rate = ifelse(Total_Clicks > 0, Total_Bookings / Total_Clicks, NA)  # Calculate conversion rate
  ) %>%
  arrange(desc(Avg_Click_Through_Rate))  # Sort by descending CTR


##Optimize Keywords##

# Count keyword frequency
my_data %>%
  count(Keyword, sort = TRUE)

# Top 10 keywords with the highest conversion rate
my_data %>%
  group_by(Keyword) %>%  # Group by Keyword
  summarise(
    Total_Clicks = sum(Clicks, na.rm = TRUE),  # Sum of Clicks
    Total_Bookings = sum(`Total Volume of Bookings`, na.rm = TRUE),  # Sum of bookings
    Conversion_Rate = ifelse(Total_Clicks > 0 & Total_Bookings > 10, Total_Bookings / Total_Clicks, NA)  # Calculate conversion rate
  ) %>%
  arrange(desc(Conversion_Rate)) %>%  # Sort by highest conversion rate
  head(20)  # Get top 20 keywords


#######Cleaning the data for further analysis
#Looking into the variables with character type data.
table(my_data$`Publisher ID`)
table(my_data$`Publisher Name`)
table(my_data$`Keyword ID`)
table(my_data$`Keyword Group`)
table(my_data$`Category`)
table(my_data$`Bid Strategy`)
table(my_data$`Keyword Type`)
table(my_data$`Status`)

#Removing useless data for analysis
new_data <- my_data[, -c(which(names(my_data) == "Publisher ID"),   # Removed because the same as "Publisher Name" 
                         which(names(my_data) == "Keyword ID"),     # Removed because code assigned to each keyword has no statistical meaning
                         which(names(my_data) == "Keyword Group"),  # Removed because the grouping was mostly based on "Campaign"
                         which(names(my_data) == "Category"),       # Removed because most of them were based on keyword
                         which(names(my_data) == "Bid Strategy"),   # Removed because ....
                         which(names(my_data) == "Keyword Type"),   # All unassigned
                         which(names(my_data) == "Avg. Pos."))]     # Removed because.....

print(head(new_data))   #for checking outcome

# Checking if there is N/A in the data
colSums(is.na(new_data))                                            #counts missing value on each variable
sapply(new_data, function(x) sum(x == "N/A", na.rm = TRUE))         #Counts "N/A" on each variable

# Creating dummy variables for 'Publisher Name' and 'Status'
new_data_dummy <- dummy_cols(new_data, select_columns = c("Publisher Name", "Status"))  # Not removing any variable on this code

#Rescaling "Match Type" from 1-5
new_data_dummy[which(new_data_dummy$`Match Type` == "Exact")   , c("Match_power")] <- "5"
new_data_dummy[which(new_data_dummy$`Match Type` == "Advanced"), c("Match_power")] <- "4"
new_data_dummy[which(new_data_dummy$`Match Type` == "Standard"), c("Match_power")] <- "3"
new_data_dummy[which(new_data_dummy$`Match Type` == 'Broad')   , c("Match_power")] <- "2"
new_data_dummy[which(new_data_dummy$`Match Type` == 'N/A')     , c("Match_power")] <- "1"
new_data_dummy$Match_power <- as.numeric(new_data_dummy$Match_power)


### Finding variable has most impact within Business Success through simple linear regression.
# Business success is Booking.
# Selecting rows where Total Volume of Bookings => 1
success_data <- new_data_dummy %>%
  filter(`Total Volume of Bookings` >= 1)

# View the filtered data
head(success_data)

colnames(success_data) # Printing variable names

my_linear_success <- lm(`Total Volume of Bookings` ~ Clicks+  # LNS on bookings above and equal to 1
                          `Avg. Cost per Click`+
                          Impressions+
                          `Publisher Name_Google - US`+
                          `Publisher Name_MSN - Global`+      # MSN US was removed to avoid multicolinearity caused by dummies.
                          `Publisher Name_Google - Global`+
                          `Publisher Name_Overture - Global`+
                          `Publisher Name_Overture - US`+
                          `Publisher Name_Yahoo - US`+
                          `Status_Live`+
                          `Status_Paused`+
                          `Status_Sent`+
                          `Status_Unavailable`+
                          `Match_power`,
                        data=success_data)
summary(my_linear_success)

my_linear_all <- lm(`Total Volume of Bookings` ~ Clicks+ # LNS on all data, including 0 bookings
                      Impressions+
                      `Publisher Name_Google - US`+
                      `Publisher Name_MSN - Global`+
                      `Publisher Name_Google - Global`+
                      `Publisher Name_Overture - Global`+
                      `Publisher Name_Overture - US`+
                      `Publisher Name_Yahoo - US`+
                      `Status_Live`+
                      `Status_Paused`+
                      `Status_Sent`+
                      `Status_Unavailable`+
                      `Match_power`,
                    data=new_data_dummy)
summary(my_linear_all)

#Normalizing____________________________________________________________________
normalize <- function(my_var){
  min_max <- (my_var-min(my_var))/(max(my_var)-min(my_var))
  return(min_max)
}

#Normalizing all variables
new_data_dummy$search_engine_bid_norm <- normalize(my_var = new_data_dummy$`Search Engine Bid`)
new_data_dummy$clicks_norm <- normalize(my_var = new_data_dummy$Clicks)
new_data_dummy$click_charge_norm <- normalize(my_var=new_data_dummy$`Click Charges`)
new_data_dummy$avg_cost_per_click_norm <- normalize(my_var=new_data_dummy$`Avg. Cost per Click`)
new_data_dummy$impressions_norm <- normalize(my_var = new_data_dummy$Impressions)
new_data_dummy$engine_ctr_norm <- normalize(my_var = new_data_dummy$`Engine Click Thru %`)
new_data_dummy$trans_conv_norm <- normalize(my_var=new_data_dummy$`Trans. Conv. %`)
new_data_dummy$amount_norm <- normalize(my_var = new_data_dummy$Amount)
new_data_dummy$total_cost_norm <- normalize(my_var = new_data_dummy$`Total Cost`)
new_data_dummy$total_bookings_norm <- normalize(my_var = new_data_dummy$`Total Volume of Bookings`)
new_data_dummy$google_us_norm <- normalize(my_var=new_data_dummy$`Publisher Name_Google - US`)
new_data_dummy$msn_global_norm <- normalize(my_var = new_data_dummy$`Publisher Name_MSN - Global`)
new_data_dummy$google_global_norm <- normalize(my_var=new_data_dummy$`Publisher Name_Google - Global`)
new_data_dummy$overture_global_norm <- normalize(my_var = new_data_dummy$`Publisher Name_Overture - Global`)
new_data_dummy$overture_us_norm <- normalize(my_var = new_data_dummy$`Publisher Name_Overture - US`)
new_data_dummy$yahoo_us_norm <- normalize(my_var=new_data_dummy$`Publisher Name_Yahoo - US`)
new_data_dummy$live_norm <- normalize(my_var = new_data_dummy$Status_Live)
new_data_dummy$paused_norm <- normalize(my_var = new_data_dummy$Status_Paused)
new_data_dummy$sent_norm <- normalize(my_var = new_data_dummy$Status_Sent)
new_data_dummy$unavailable_norm <- normalize(my_var = new_data_dummy$Status_Unavailable)
new_data_dummy$match_power_norm  <- normalize(my_var = new_data_dummy$Match_power)

#Model Building_________________________________________________________________

#Determining the Business Success based on Total Bookings, where x>0 = success
table(new_data_dummy$`Total Volume of Bookings`)
booked <- new_data_dummy[which(new_data_dummy$`Total Volume of Bookings`>0),c("booked")]<-"1"
booked <- new_data_dummy[which(new_data_dummy$`Total Volume of Bookings`==0),c("booked")] <-"0"
new_data_dummy$booked <- as.numeric(new_data_dummy$booked)
summary(new_data_dummy)


#Splitting and testing nd_d = New Data Dummy

# Perform stratified sampling based on the 'booked' column
# Stratification is necessary because 'booked = 0' occurs significantly more frequently than 'booked = 1'
new_data_dummy_strat <- stratified(as.data.frame(new_data_dummy), 
                                   group = "booked",
                                   size = 0.8, 
                                   bothSets = TRUE)  # Returns both training and testing datasets

# Extract the training dataset (80% of the original dataset)
new_data_dummy_train <- new_data_dummy_strat$SAMP1

# Extract the testing dataset (20% of the original dataset)
new_data_dummy_test <- new_data_dummy_strat$SAMP2

#Logistic Regression
af_logist <- glm(booked~ Clicks+ # LNS including 0 bookings
                   Impressions+
                   `Avg. Cost per Click`+
                   `Publisher Name_Google - US`+
                   `Publisher Name_MSN - Global`+
                   `Publisher Name_Google - Global`+
                   `Publisher Name_Overture - Global`+
                   `Publisher Name_Overture - US`+
                   `Publisher Name_Yahoo - US`+
                   `Status_Live`+
                   `Status_Paused`+
                   `Status_Sent`+
                   `Status_Unavailable`+
                   `Match_power`,
                 data=new_data_dummy_train,
                 family="binomial") #This uses the data with units
summary(af_logist)


exp(7.713e-03)  - 1  # For every additional unit of clicks, the odds of booking will increase by 0.77%.
exp(-3.646e-06) - 1  # For every additional unit of impressions, the odds of booking will decrease by 0.00036%.
exp(-2.490e-01) - 1  # For every additional unit increase in the average cost per click, the odds of booking will decrease by 22.04%.
exp(1.714e+00)  - 1  # Being associated with Google US increases the odds of booking by 355.11%.
exp(2.170e+00)  - 1  # Being associated with MSN Global increases the odds of booking by 675.83%.
exp(2.218e+00)  - 1  # Being associated with Google Global increases the odds of booking by 718.89% .
exp(2.818e+00)  - 1  # Being associated with Overture Global increases the odds of booking by 1,574.33% .
exp(2.070e+00)  - 1  # Having a keyword status of "Live" increases the odds of booking by 692.48%.

#Designing unitless mode:
af_logist_norm <- glm(booked~ clicks_norm+ # LNS including 0 bookings
                        impressions_norm+
                        avg_cost_per_click_norm+
                        google_us_norm+
                        msn_global_norm+
                        google_global_norm+
                        overture_global_norm+
                        overture_us_norm+
                        yahoo_us_norm+
                        live_norm+
                        paused_norm+
                        sent_norm+
                        unavailable_norm+
                        match_power_norm,
                      data=new_data_dummy_train,
                      family="binomial") #This uses the data with units
summary(af_logist_norm)

# One-unit increase in Clicks causes significant increase in the the odds of booking.
# On the other hand Impressions decreases the odds of booking significantly, which may indicates that more impressions without clicks could indicate low ad relevance.
# The company's global presence is higher than the US. Google, Overture, and MSN shows better impact compared to others.
# This shows that the company should focus on international travel.
# Furthermore, higher  avg cost per click  reduces the odds of booking, which illustrates higher CPC does not bringing customers/conversions/.

#implementing the confusion matrix
my_prediction <- predict(af_logist,new_data_dummy_test, type="response") #predicting on test data

confusionMatrix(data= as.factor(as.numeric(my_prediction > 0.5)), 
                reference= as.factor(as.numeric(new_data_dummy_test$booked)))

#The accuracy of the model based on test is 93.4%, which is great.

#coding a gini

my_tree<- rpart(booked~ Clicks+
                  Impressions+
                  `Avg. Cost per Click`+
                  `Publisher Name_Google - US`+
                  `Publisher Name_MSN - Global`+
                  `Publisher Name_Google - Global`+
                  `Publisher Name_Overture - Global`+
                  `Publisher Name_Overture - US`+
                  `Publisher Name_Yahoo - US`+
                  `Status_Live`+
                  `Status_Paused`+
                  `Status_Sent`+
                  `Status_Unavailable`+
                  `Match_power`,
                data=new_data_dummy_train,
                method="class",
                cp=0.01)
rpart.plot(my_tree)


#designing confusion matrix for the tree
tree_predict <- predict(my_tree, new_data_dummy_test, type = "prob")
confusionMatrix(data= as.factor(as.numeric(tree_predict[,2] >0.5)),
                reference= as.factor(as.numeric(new_data_dummy_test$booked)))

#The accuracy of the gini tree index based on test is 93.35%, which is closer to the logistic regression model.

# comparing both models logit and tree on one framework through AUC ROC analysis.
logit_rocr<- prediction(my_prediction,new_data_dummy_test$booked)
perf_logit <- performance(logit_rocr, "tpr", "fpr")
plot(perf_logit, lty=3, lwd=3)

tree_rocr<- prediction(tree_predict[,2],new_data_dummy_test$booked)
perf_tree <- performance(tree_rocr, "tpr", "fpr")
plot(perf_tree, lty=3, lwd=3, col="green", add=TRUE)

#Based on how the logistic regression true positive result curved more on the left side, we conclude that logistic regression model is better.



