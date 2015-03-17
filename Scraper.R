#Scraping Top Seller Reviews from Amazon

require(RCurl)
require(XML)
require(xlsx)


catg<-c("Books","Movies-Tv","Electronics", "toys-and-games", "apparel")

for (cat in 1:length(catg))
{
  productsFile<-paste0("Amazon_BestSeller_",catg[cat],"_20.csv")
  amazon_products<-read.csv(productsFile, header=F, stringsAsFactors=F)
  
  reviewsFile<-paste0("Amazon_BestSeller_Top20_",catg[cat],"_Reviews100.xlsx")
  wb <- createWorkbook() #create a new workbook
  
  for (i in 1:nrow(amazon_products))
  {
      reviews_100<-character(0)
      
      product_name<-unlist(strsplit(amazon_products[i,2], "/"))[4]
      product_code<-unlist(strsplit(amazon_products[i,2], "/"))[6]
      
      #cat(product_name," ~ ", product_code, "\n")
      #reviews_file<-paste0("Amazon_Reviews_",product_name,".rds")
      #reviews_file<-paste0("Amazon_Reviews_",product_name,".csv")
      
      #Create a new sheet for each product in category
      sheet <- createSheet(wb, sheetName=paste0(product_name,"_",product_code))
      
      for (page_id in 1:10)
      {
          URL<-paste0("http://www.amazon.com/",product_name,"/product-reviews/",product_code,"/ref=cm_cr_pr_top_link_",
                      page_id,"?ie=UTF8&pageNumber=",page_id,"&showViewpoints=0&sortBy=byRankDescending")
          cat("\n",URL)
          Source<-getURL(URL,encoding="UTF-8")
          
          PARSED<-htmlParse(Source)
          xpathSApply(PARSED, "//*[@id='productReviews']",xmlValue)
          
          reviews<-xpathSApply(PARSED, "//*[@class='reviewText']",xmlValue)
          reviews_100<-c(reviews_100,reviews)
      }
      reviews_df<-data.frame(cbind(rep(product_name,length(reviews_100)), rep(product_code, length(reviews_100)), 
                                   unlist(reviews_100))
                             
      colnames(reviews_df)<-c("Product Name", "Product Code", "Top 100 Reviews")
      #saveRDS(reviews_df,reviews_file)
      write.csv(reviews_df, file=reviews_file, row.names=F)
      addDataFrame(reviews_df, sheet)   
  }
  saveWorkbook(wb, reviewsFile)
}
