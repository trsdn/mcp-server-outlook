# CLI Command Reference

> Generated from `outlookcli --help`. Do not edit manually.

## application

```
DESCRIPTION:
Application operations

USAGE:
    outlookcli application <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                             Prints help          
                                                           information          
        --include-active-context <INCLUDEACTIVECONTEXT>    IncludeActiveContext 
    -o, --output <PATH>                                    Write output to file 
                                                           instead of stdout.   
                                                           For image results,   
                                                           decodes and saves as 
                                                           binary file          
```

## attachment

```
DESCRIPTION:
Attachment operations

USAGE:
    outlookcli attachment <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                            Prints help           
                                                          information           
        --mail-entry-id <MAILENTRYID>                     MailEntryId           
        --store-id <STOREID>                              StoreId               
        --use-active-mail <USEACTIVEMAIL>                 UseActiveMail         
        --destination-directory <DESTINATIONDIRECTORY>    (required for: save)  
        --attachment-index <ATTACHMENTINDEX>              (required for: remove)
        --overwrite <OVERWRITE>                           Overwrite             
        --file-path <FILEPATH>                            (required for: add)   
    -o, --output <PATH>                                   Write output to file  
                                                          instead of stdout. For
                                                          image results, decodes
                                                          and saves as binary   
                                                          file                  
```

## calendar

```
DESCRIPTION:
Calendar operations

USAGE:
    outlookcli calendar <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                             Prints help          
                                                           information          
        --folder <FOLDER>                                  Folder               
        --start <START>                                    (required for:       
                                                           create-appointment)  
        --end-time <ENDTIME>                               (required for:       
                                                           create-appointment)  
        --max-count <MAXCOUNT>                             MaxCount             
        --include-body-preview <INCLUDEBODYPREVIEW>        IncludeBodyPreview   
        --entry-id <ENTRYID>                               EntryId              
        --store-id <STOREID>                               StoreId              
        --use-active-appointment <USEACTIVEAPPOINTMENT>    UseActiveAppointment 
        --subject <SUBJECT>                                (required for:       
                                                           create-appointment)  
        --location <LOCATION>                              Location             
        --body <BODY>                                      Body                 
        --all-day <ALLDAY>                                 AllDay               
        --display <DISPLAY>                                Display              
    -o, --output <PATH>                                    Write output to file 
                                                           instead of stdout.   
                                                           For image results,   
                                                           decodes and saves as 
                                                           binary file          
```

## folder

```
DESCRIPTION:
Folder operations

USAGE:
    outlookcli folder <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                       Prints help information    
        --include-item-counts <INCLUDEITEMCOUNTS>    IncludeItemCounts          
        --parent-folder <PARENTFOLDER>               ParentFolder               
        --folder <FOLDER>                            Folder                     
        --include-item-count <INCLUDEITEMCOUNT>      IncludeItemCount           
        --max-count <MAXCOUNT>                       MaxCount                   
        --include-preview <INCLUDEPREVIEW>           IncludePreview             
    -o, --output <PATH>                              Write output to file       
                                                     instead of stdout. For     
                                                     image results, decodes and 
                                                     saves as binary file       
```

## mail

```
DESCRIPTION:
Mail operations

USAGE:
    outlookcli mail <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                         Prints help information  
        --entry-id <ENTRYID>                           EntryId                  
        --store-id <STOREID>                           StoreId                  
        --use-active-mail <USEACTIVEMAIL>              UseActiveMail            
        --folder <FOLDER>                              Folder                   
        --max-count <MAXCOUNT>                         MaxCount                 
        --unread-only <UNREADONLY>                     UnreadOnly               
        --include-body-preview <INCLUDEBODYPREVIEW>    IncludeBodyPreview       
        --query <QUERY>                                (required for: search)   
        --recipient-to <RECIPIENTTO>                   RecipientTo              
        --cc <CC>                                      Cc                       
        --bcc <BCC>                                    Bcc                      
        --subject <SUBJECT>                            (required for:           
                                                       set-subject)             
        --body <BODY>                                  (required for: set-body) 
        --display <DISPLAY>                            Display                  
        --confirm <CONFIRM>                            Confirm                  
        --operation-id <OPERATIONID>                   OperationId              
        --target-folder <TARGETFOLDER>                 (required for: move)     
        --is-read <ISREAD>                             (required for:           
                                                       set-read-state)          
        --categories <CATEGORIES>                      Categories               
    -o, --output <PATH>                                Write output to file     
                                                       instead of stdout. For   
                                                       image results, decodes   
                                                       and saves as binary file 
```
