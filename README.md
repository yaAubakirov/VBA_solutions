# VBA_solutions
Here I store my VBA macros


### 1. queryAuto macro

  At work my team is resposible for filling one weekly report. Since team is working in a construction field they send TQ (Technical Query) and recieves the answers. So this answers should be shown in special sheet. 
  
  The main problem is that there are different construction modules and each module divided by many DPs (delivery pack). So each TQ affects to single DP, but could affects to many ones.
  
  The module name looks like this 1-TCD-001. DP is written like this DP1. So affecting DPs should be written in a row in a different cells as follows 1D1DP1, 1D1DP2 etc.
  
  TQ numbers looks like this RHIQ-TPK-1-TCD-001-0001. RHIQ is contractor's code, TPK - short name of sub-contractor who raised TQ, 1-TCD-001 - is module name, 0001 is TQ number. But also we have General TQs which ones affect to whole project. So the number is like this RHIQ-TPK-General-0002. For this TQs no need to write all DPs, but there is a special column which one shows if this TQ is General.
  
  There is a special person who is responsible for TQs. He organize them in a special table. The problem is he writes DPs in a single cell. So if I need to fill them as Contractor requests I need to extract DPs one by one, concat them with module name and insert it to dedicated cell.
  
  So first macro checks if this TQ is General. If it is General TQ it writes Yes to dedicated columns, in another result writes No. After that for non-General TQs it takes short name of module concats it with first DP in a cell and inserts to dedicated cell. And so on till the end
  
  This what macro does. Example is attached

### 2. submissionsFill macro

One guy asked me to creat macro which one takes data from one file and fills another based on conditions.

As I explaned before there are modules and DPs. DPs are submitted for checking and contractor assigns status for DP. Code 1 - couldn't be approved, need to correct. Code 2 - approved but there are small corrections. Code 3 - approved.

So based on marged module and DP string need to find statuses for DP and fill statuses in a row in a chronological order. If DP wasn't send need to write "Not submitted". If it is submitted but it is still on checking need to write "Under Review"

As an example there are files 'table_to_fill.xlsm' and 'table_to_parse.xlsx'. Open first file, check VBA and run module, select file 'table_to_parse.xlsx' and enjoy
