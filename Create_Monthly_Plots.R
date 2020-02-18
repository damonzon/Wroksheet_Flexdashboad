# Plots of Total A by Month (2012-2017)
suppressMessages(library(xts)) 
TA_By_Month <-function(number = "January",cell="B2:B3"){
    read_excel_allsheets <- function(filename, tibble = FALSE) {
        sheets <- readxl::excel_sheets(filename)
        x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X,range = cell ))
        if(!tibble) x <- lapply(x, as.data.frame)
        names(x) <- sheets
        x
    }
    TG <- read_excel_allsheets("Details_Monthly.xls")
    TG <- as.data.frame(TG)
    TG <- TG[c(6:1)]
    TG2 <- data.frame(NULL)
    for (i in 1:6){
        TG2 <- c(TG2,TG[,i])
    }
    month <<- as.vector(unlist(TG2))
    Which_Month <<- number
}

# January
TA_By_Month("January","B2:B3")
Which_Month
month

# Transform data to time series class
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))
# month_model <- lm(month_ts~time(month_ts))
# month_model$coefficients[2]

png(filename="A_January.jpg")
plot(month_ts, main = "January Total A: 2012-2017\nRed = Trend Line -- Green = Median",
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()


Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_February.jpg")
# plot.new()
#plot(month_ts, main = "February Total A: 2012-2017\nRed = Trend Line -- Green = Median",
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# March
TA_By_Month("March","B4:B5")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_March.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# April
TA_By_Month("April","B5:B6")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_April.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()
# May
TA_By_Month("May","B6:B7")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_May.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# June
TA_By_Month("June","B7:B8")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_June.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()
# July
TA_By_Month("July","B8:B9")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_July.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# August
TA_By_Month("August","B9:B10")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_August.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# September
TA_By_Month("September","B10:B11")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_September.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# October
TA_By_Month("October","B11:B12")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_October.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# NOvember
TA_By_Month("November","B12:B13")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_November.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()

# December
TA_By_Month("December","B13:B14")
Which_Month
month
month_ts <- ts(month, frequency=1, start=c(2012,1), end = c(2017,1))

png(filename="A_December.jpg")
plot(month_ts, main = paste(Which_Month,"Total A: 2012-2017\nRed = Trend Line -- Green = Median"),
     ylim = c(200,1200),
     xlab = "Year", ylab = "Count")
abline(v = 2012:2017, col = "blue")
abline(lm(month_ts~time(month_ts)),col = "red")
abline(h = median(month_ts), col = "darkgreen")
dev.off()


