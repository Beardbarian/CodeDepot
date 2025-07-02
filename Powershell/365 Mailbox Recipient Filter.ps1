#Run as a single line
Set-DynamicDistributionGroup -Identity 'DL Name' 
-RecipientFilter {((RecipientType -eq 'UserMailbox') 
-and -not(RecipientTypeDetailsValue -eq 'SharedMailbox' 
-or RecipientTypeDetailsValue -eq 'GuestMailUser' 
-or Name -like 'Invoices' 
-or Name -like 'peter' 
-or Name -like 'steve' 
-or Name -like 'support' 
-or Name -like 'social' 
-or Name -like 'charlie' 
-or Name -like 'events' 
-or Name -like 'info' 
-or Name -like 'gatherall'))}
