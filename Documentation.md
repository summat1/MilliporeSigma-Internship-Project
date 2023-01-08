# Legacy Tissue Data Project
### Shivesh Ummat - Bioinformatics Intern at MilliporeSigma 
#### 12/16/2022

Over a 12 week summer internship and part-time work throughout the autumn, I have worked on the **Legacy Tissue Data Project**. The purpose of this project is to organize 3 Excel sheets of data: **Closet Blocks**, **Customer Blocks**, and **Active Inventory**. Closet and Customer Blocks contain information about the shipment and receiving of tissue blocks. Some of these are qualified and have a place in Active Inventory, while some are unqualified and remain in storage. Closet and Customer Blocks may be herein referred to as **Sourcing Data** and Active Inventory as **QTI** (Qualified Tissue Inventory). 

This document serves as a documentation for the steps I took in this project, the results I found, and the future implications for my work. It will follow the steps I performed in chronological order, explaining the methodology and purpose behind each part. It will be attached with the frozen versions of the data files that I used, the code files in Python and R, and the outputted data files after data wrangling. All the source code is commented as well for further documentation.

The end goal of this project is to integrate the legacy tissue data into a **Laboratory Information Management System (LIMS)**. Significant progress has been made towards this goal, although a few key challenges remain. 

### Closet Blocks
#### Overview of Data
1. The **Closet Blocks** spreadsheet contains data on tissue samples received from surgical procedures. This Excel file is organized into many sheets, where each sheet represents the tissue type of the blocks received. This data is a source for blocks in active inventory. A majority of these blocks are unqualified. 

2. One key note is the discrepancy in the formatting of tissue block IDs (CM Number below) in this data file. As this is the key identifier for the tissue blocks, I spent time editing the file manually and used this for my analysis. I highlighted every edit I made according to the following color code: yellow highlight for simple edits (removing spaces, obvious typos) and orange for edits that need more review but were necessary for the code to work (changing syntax, typing in replacements for block IDs that were unreadable). This file is called *Closet Blocks Highlighted Edit.xlsx*. 

#### Data Columns:
- Surgical Number
    - Number assigned by the surgical department from the donor
- Specimen
    - Tissue type/location
- Procedure
    - Surgical procedure the tissue came from
- Total Number of Blocks
    - How many were sent from the provider
- Number of Blocks used
    - How many could be used by CM
- CM Number
    - The corresponding CM numbers to the blocks used
    - If the number is PST-78A-G for example, this is 7 blocks from A to G
- Initials/H&E Date
    - Reception of blocks information (date and receiver initials)
- Comment
    - Any additional info, including if the samples were discarded
- \# Blocks Discarded
- Location
    - Storage location

#### Goal of Data Manipulation
1. The dataset should be combined into one sheet. 
2. The vendor/hospital should be identified via the surgical number in a new column called "Source." 
3. Rows should be split apart so that each row only represents one tissue block.
4. Unqualified and qualified blocks should be separated.

#### Data Manipulation Process
1. The dataset was combined from multiple sheets into one sheet. In order to preserve the information contained in the sheet naming convention, an extra column called "Category" was added with the sheet name. 

*Export: "Simply Combined Closet Blocks.xlsx"*

2. Sourcing information added. Any surgical numbers starting with "P" or "D" (lowercase or capital) were assigned to Princeton Baptist. Any surgical numbers containing "LV" (lowercase or capital) were assigned to Lakeview Hospital. Any surgical numbers containing "ML" or "MF" (lowercase or capital) were assigned to NatPkMedCtr AMI HotSprings. If none of these are true, the source is Unknown.
3. In order to split rows, we must look at the features "Total Number of Blocks" and "Number of Blocks used." The difference (if positive) between these two numbers was extracted as unqualified blocks. These were turned into new rows. The remaining qualified blocks were then split into individual rows by using robust methods fully documented in the source code. These methods split up the Block IDs and assign one unique block ID to each unique block in separate rows. After this, the "Total Number of Blocks" was reassigned to 1 and "Number of Blocks used" was reassigned to 1 or 0 depending on if the block was qualified or not. At this point, the data was exported. 
*Export: "Expanded and Combined Closet Blocks.xlsx."*

4. Qualified blocks ("Number of Blocks Used" = 1) were categorized into a separate place as unqualified blocks ("Number of Blocks Used" = 0). 

*If the "Total Number of Blocks" is LESS THAN than the "Number of Blocks used," this is paradoxical and indicates an error. Export: "Error Rows.xlsx"*

*Qualified blocks were exported. Export: "Cleaned Closet Blocks.xlsx"*

*Unqualified blocks were exported. Export: "Unqualified Closet Blocks.xlsx"*
### Customer Blocks
#### Overview of Data
1. The **Customer Blocks** spreadsheet contains information on the blocks we received from customers. Each sheet in this Excel file represents a different vendor but contains all tissue types. This data is a source for the blocks in active inventory.

2. One key note is the discrepancy in the formatting of tissue block IDs (Keepers: CM ID# below) in this data file. As this is the key identifier for the tissue blocks, I spent time editing the file manually and used this for my analysis. I highlighted every edit I made according to the following color code: yellow highlight for simple edits (removing spaces, obvious typos) and orange for edits that need more review but were necessary for the code to work (changing syntax, typing in replacements for block IDs that were unreadable). This file is called *Customer Blocks Highlighted Edit.xlsx*. 

#### Data Columns:
- Receiving Date
    - MM/DD/YYYY
- Sales Rep
- Expectation
    - Expected pathology
- \# of Blocks
    - Number of blocks received
- Received by
    - Name of who received the samples	
- Tissue type	
- Keepers: CM ID#
    - For samples kept, this indicates the Cell Marque Block ID it became. 
- Return, Discard, or Credit date & Initial
    - Indicates what happened to the samples
- Comments/Customer Block ID #	
    - The ID used by the customer
- Demographic Info	
- Location
    - Storage location	

#### Goal of Data Manipulation
1. The dataset should be combined into one sheet.  
2. Rows should be split apart so that each row only represents one tissue block.
3. Unqualified and qualified blocks should be separated.
#### Data Manipulation Process
1. Each sheet in the Excel file was combined into one sheet. Since the sheet name was used to identify the vendor, a new column called "Source" was created with the vendor name, retaining this information. 
2. It appeared that certain data was not input but assumed to apply for an entire shipment. For example, if many blocks were received on the same day and input over multiple rows, only the first row might contain info such as "Sales Rep" or "Receiving Date." So, if not present, this info was filled in for shipments that came in on the same day. 
3. The qualified blocks were split into individual rows by using robust methods fully documented in the source code. These methods split up the Block IDs and assign one unique block ID to each unique block in separate rows. 
4. The qualified and unqualified blocks were split and exported separately. 

*Export: "Cleaned Customer Blocks.xlsx"*

*Export: "Unqualified Customer Blocks.xlsx"*

5. Note: Rows containing both unqualified and qualified blocks were not split. For example, if a shipment contained 10 blocks and 3 were qualified, this was only split into 3 rows (for each qualified block) not 10. We should have 10, since each row should represent one block. This means that some unqualified blocks have been "lost in translation." This is something I didn't get a chance to work on.
### Consolidation and Merging
- At this point, both Customer and Closet blocks had been fully cleaned. In order to move forward, these datasets had to be compiled and joined with the Active Inventory. 
#### Overview of Data
- From the previous steps, we have created two cleaned datasets: *Cleaned Closet Blocks.xlsx* and *Cleaned Customer Blocks.xlsx*. We will also be working with a frozen version of the Active Inventory called *Qualified Tissue Inventory_FROZEN 6.27.22.xlsx*. 
#### Goal of Data Manipulation
- From here, we need to join these files together. We want to match data that corresponds to the same block. For example, block "AA-1A" might have corresponding data in both QTI and Customer or Closet blocks. 

- Since we can't be sure which of those two files the sourcing data will be found in, we can join the qualified Customer and Closet blocks together first in an intermediary step. This **Consolidation** will simply add all the rows from both spreadsheets into one source dataset. 

- After that, we will have to **merge** the source dataset with the QTI dataset. This will add source data to each row in QTI so we know where each block came from.
#### Data Manipulation Process
- The source dataset was created by adding all the rows from both Closet and Customer blocks into one spreadsheet: **Sourcing Data**. Each qualified block has one row in this spreadsheet. There are 16,203 rows, and therefore **16,203 qualified blocks found in Closet and Customer Blocks**. However, since two spreadsheets were combined, this effectively doubles the number of columns. The columns: "Source" and "Block ID #" are the ones that overlap between the two spreadsheets and therefore are not duplicated.

*Export: "All Sourcing Data - Qualified Only.xlsx"*

- The QTI was also prepped for merging, although this was trivial as it is very organized already. 

*Export: "Cleaned Qualified Tissue Inventory.xlsx"*

- From here, these two spreadsheets were **merged together** into the final dataset. The logic is as follows: 
If a certain "Block ID #" in QTI is found in the "Block ID #" column in the sourcing data spreadsheet, append the data contained in the sourcing file to the QTI.

*Export: "First Merge.xlsx"*

- One challenge immediately presented was that some blocks have changed their Block ID #. So, matching was also performed between "Previous Block ID" in QTI to "Block ID #" in the sourcing data. 

*Export: "Second Merge.xlsx"*

- Both of these merges were then grouped into a fully joined dataset. 

*Export: "QTI + Source Data Merged.xlsx"*

- There were **16,203 qualified blocks found from the sourcing spreadsheets** and **16,312 qualified blocks in QTI**. We expected that these would merge one-to-one and nearly all of the QTI blocks would now have source data appended to them. However, this is not how the merge actually happened. Only **9872 blocks, or 60.5% of blocks** in QTI were matched with sourcing data. Possible reasons for this will be explored in the discussion. Additionally, we will explore tactics to increase this percentage.
### Final Data Manipulation
- At this point, one **master file** has been assembled. Now some work must be done to organize this dataset and source more blocks (close the 40% gap of unsourced blocks). 
#### State of Data
- The 3 datasets have been cleaned, consolidated, and merged together. This master file (*QTI + Source Data Merged*) will be the dataset going forward.
#### Goal of Data Manipulation
- Additional data manipulation must be done to format this data into LIMS-ready state. First, dates should be formatted as "DD-Mmm-YY." Next, columns named "Dx change Yes/No" and "Source Type" should be added to characterize each block further. Lastly, preliminary work can be done to expand the more complicated columns in QTI such as the antibody scores.
#### Data Manipulation Process
1. The very first step was to complete the dataset. Since the merge only accounted for 9872 tissue blocks, the rest were added in from QTI. This brings us to a total of 16,307 qualified blocks. 
2. Data wrangling methods were applied to format the date in the correct way. 
3. Calculated column "Source Type" was added depending on which source information was appended to the block (closet vs customer source). A column called "Dx Change? Yes/No" was added. Yes if there is a previous block ID, no if not.
4. Possibly the most important step in this section was establishing the relationship between parent and child blocks. We established that two blocks with the **same patient number** (see Block ID # data formatting) should have the same source. Therefore, I created and applied a robust method to split apart a Block ID # into its component parts and used this to identify which blocks came from the same patient. I then appended the source data to these related blocks. This brought the **total sourced blocks up to 11,253**. The IDs of the blocks sources found were extracted for further analysis.

*Export: "No Source.xlsx"*

5. Exploratory work with positive antibody scores was done as a proof-of-concept that we can extract the meaningful information from these scores if formatted properly.

*Export: "Positive Scores.xlsx"*

6. After this process was complete, the final dataset was exported. 

*Export: "Master File of All Data.xlsx"*

Note: The Block ID Numbers are used as indices.
### Analytics and Manipulation in R
- This concludes the in-depth technical portion of the project. Analytics were done in R to understand some of the high-level results of this project.
#### State of Data
- The Master File was used for this analysis.
#### Goal of Data Manipulation
1. Add a "Tissue Type" column
2. Characterize the blocks that do not have sourcing data.
#### Data Manipulation Process
1. Tissue Type column was added by using the abbreviations list from Histology. This could have been done in Python but was giving trouble so I did it in R. 
2. This allowed me to perform high level analytics on the non-sourced blocks in order to understand why we cannot find a source.

*Export: "Final Master File.csv"*

*Export: "Non-Sourced Statistics.csv"*
### Summary
- Throughout this process, 3 frozen data sheets containing information about the same tissue blocks were consolidated into one dataset. Out of 16,312 qualified blocks in Active Inventory, 11,253 or 69% were found to have source data. 

- To maintain a steady record of data analysis and quality control, the state of the data was analyzed and exported at various regular intervals. Both Python and R were used for data cleaning/manipulation and analytics, respectively. 

- 6 separate scripts were written and commented to track the data and its alterations along this process. 

- The end result is the Final Master File of all the Cell Marque Tissue Blocks Data.
### Discussion
While we hoped for more than 90% of QTI to match with sourcing data, we only sourced 70% of blocks. The missing 30% of sourcing data is almost certainly due a **combination of many reasons**.
- 16,203 qualified blocks were found from the two sourcing spreadsheets. If only 9,872 of these blocks matched to QTI, what is happening to the rest? Are they qualified? They might not be in Active Inventory, but they should belong somewhere. It seems this number could be inflated. This may be because many of the block IDs in the sourcing spreadsheets were ~~struckthrough~~. Cleaning needs to be done on these spreadsheets to understand which blocks were actually qualified. 
- Are there more relationships that can tag blocks together? Currently, source data is appended to blocks that match directly with the Block ID, or were previously named as that Block ID, or came from the same patient. Can we derive more relationships that will allow us to source more blocks?
- Tissue blocks are often being recharacterized and renamed. The "Previous Block ID" column might have multiple entries. This will prevent matching. We may need to clean up this column in order to improve sourcing results.
- Analysis in R shows that a high percentage of blocks with certain tissue types like Urothelial Carcinoma and Tonsil, among others, are not sourced. Is there a reason for this? How might we look back into the Closet and Customer blocks to find where these blocks are coming from?
### Future Directions and Implications
There are many important future directions for this work. 
1. **Expanding/extracting data from certain columns**. 
Some of the features in this dataset are quite complex. In order to glean categorizable information from these columns, robust methods and algorithms must be developed to scan through the data and grab the key elements. This can be true for: "Negative/Positive Antibody Scores," "Diagnostic Comments," and "Other Comments" (just from QTI). Some of the columns from the sourcing data are also like this, such as "Initials/H&E Date." There is no clear formatting for some of these columns and they may be very difficult to glean information from.
2. **Implementing a consistent syntax structure**.
In order to address this issue, we can implement a syntax structure that makes the data more organized and consistent. This would help with gleaning information from the complex data columns. I have created a Syntax Guide for the Rocklin Site. However, it still may need improvement and testing before it can be implemented.
3. **Database merging with SQL**. 
The database-style merging is extremely limited in Python's Pandas library. I found it challenging to perform merges and concatenations especially because we need to match with "Block ID #" and "Previous Block ID" in QTI. SQL would be better used for this purpose. However, I have not learned this language and was unable to use it for this project.
4. **Continuing to close the unsourced gap**. 30% of blocks in QTI still need to be matched to a source. Some of the reasons in _Discussion_ may be the key to this step. 
5. **Block Status**. The sourcing data may list a blocks that have been discarded. We must factor this into our analysis. Recommendation has been made to create a column blocks called "Status" so that we know if blocks are qualified/active. unqualified and fit to be examined, or were discarded after reception. This will help track all the blocks that came into Cell Marque.
6. **Patient Identifiers**. Before LIMS integration, it has been recommended to create a "Patient Identifier" column. This will help keep track of a patient relationship to understand which blocks came from the same patient. Preliminary work into this was done by analyzing the components of an individual block ID to grab the "patient number." Further work needs to be done to map block relationships.
7. **Quality Control**. Quality in the data needs to be maintained. While the Excel file can be messy, it is the working record and contains all the correct data. LIMS integration must maintain the integrity of this data and therefore must go through quality checks.
8. **LIMS Integration**. The final step in this process is LIMS Integration. It may be a while before this is acheived. But this will streamline our operations and significantly improve decision-making both in and out of the lab. I believe this integration is necessary for an operation as large as the one at the Rocklin site. The Excel files can be difficult to index and understand without long-standing experience. A large-scale online software will perfectly help with this issue. Through all the steps outlined in this document, LIMS integration for the legacy data should be possible. Once this data moves into the software, everyone can begin to use the platform moving forward. I hope my project has started this process and faciliated the next steps.
### Acknowledgements
First and foremost, I would like to acknowledge my mentor and manager Camille Gomez. Camille introduced me to Cell Marque, facilitated my learning, and helped me push through obstacles in this project. I learned so much from her and I am so thankful she orchestrated this opportunity for me. I would also like to thank Lauren Nguyen, who welcomed me into MilliporeSigma and has supported me throughout this journey. Lauren helped make sure I maximized my experience in this role. I'd like to thank the Cell Marque R&D team for contributing to my project via presentations and meetings. Many members indirectly supported this project as well. Finally, thanks to the Randstad team for making this internship logistically possible.

Thank you,

***Shivesh Raj Ummat***