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
    -h, --help                                              Prints help
                                                            information
        --folder <FOLDER>                                   Folder
        --start <START>                                     (required for:
                                                            create-appointment)
        --end-time <ENDTIME>                                (required for:
                                                            create-appointment)
        --max-count <MAXCOUNT>                              MaxCount
        --include-body-preview <INCLUDEBODYPREVIEW>         IncludeBodyPreview
        --entry-id <ENTRYID>                                EntryId
        --store-id <STOREID>                                StoreId
        --use-active-appointment <USEACTIVEAPPOINTMENT>     UseActiveAppointment
        --subject <SUBJECT>                                 (required for:
                                                            create-appointment)
        --location <LOCATION>                               Location
        --body <BODY>                                       Body
        --all-day <ALLDAY>                                  AllDay
        --display <DISPLAY>                                 Display
        --required-attendees <REQUIREDATTENDEES>            RequiredAttendees
        --optional-attendees <OPTIONALATTENDEES>            OptionalAttendees
        --resource-attendees <RESOURCEATTENDEES>            ResourceAttendees
        --send-invitation <SENDINVITATION>                  SendInvitation
        --recurrence-type <RECURRENCETYPE>                  RecurrenceType
        --recurrence-interval <RECURRENCEINTERVAL>          RecurrenceInterval
        --recurrence-days-of-week <RECURRENCEDAYSOFWEEK>    RecurrenceDaysOfWeek
        --recurrence-count <RECURRENCECOUNT>                RecurrenceCount
        --recurrence-end-date <RECURRENCEENDDATE>           RecurrenceEndDate
        --occurrence-date <OCCURRENCEDATE>                  OccurrenceDate
        --attendees <ATTENDEES>                             (required for:
                                                            get-free-busy)
        --days <DAYS>                                       Days
        --interval-minutes <INTERVALMINUTES>                IntervalMinutes
        --file-path <FILEPATH>                              (required for:
                                                            export)
        --format <FORMAT>                                   Format string
        --overwrite <OVERWRITE>                             Overwrite
    -o, --output <PATH>                                     Write output to file
                                                            instead of stdout.
                                                            For image results,
                                                            decodes and saves as
                                                            binary file

```

## contact

```
DESCRIPTION:
Contact operations

USAGE:
    outlookcli contact <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                                   Prints help
                                                                 information
        --folder <FOLDER>                                        Folder
        --max-count <MAXCOUNT>                                   MaxCount
        --include-body-preview <INCLUDEBODYPREVIEW>              IncludeBodyPrev
                                                                 iew
        --entry-id <ENTRYID>                                     EntryId
        --store-id <STOREID>                                     StoreId
        --use-active-contact <USEACTIVECONTACT>                  UseActiveContac
                                                                 t
        --first-name <FIRSTNAME>                                 FirstName
        --last-name <LASTNAME>                                   LastName
        --company-name <COMPANYNAME>                             CompanyName
        --job-title <JOBTITLE>                                   JobTitle
        --email1-address <EMAIL1ADDRESS>                         Email1Address
        --email2-address <EMAIL2ADDRESS>                         Email2Address
        --business-telephone-number <BUSINESSTELEPHONENUMBER>    BusinessTelepho
                                                                 neNumber
        --mobile-telephone-number <MOBILETELEPHONENUMBER>        MobileTelephone
                                                                 Number
        --body <BODY>                                            Body
        --display <DISPLAY>                                      Display
    -o, --output <PATH>                                          Write output to
                                                                 file instead of
                                                                 stdout. For
                                                                 image results,
                                                                 decodes and
                                                                 saves as binary
                                                                 file

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
        --store-id <STOREID>                         StoreId
        --address <ADDRESS>                          Address
        --role <ROLE>                                Role
        --parent-folder <PARENTFOLDER>               ParentFolder
        --name <NAME>                                Name
        --folder <FOLDER>                            Folder
        --destination-folder <DESTINATIONFOLDER>     DestinationFolder
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
        --max-count <MAXCOUNT>                         How many rows to return.
                                                       The counts always
                                                       describe the full set
        --unread-only <UNREADONLY>                     UnreadOnly
        --include-body-preview <INCLUDEBODYPREVIEW>    IncludeBodyPreview
        --from-address <FROMADDRESS>                   FromAddress
        --subject-contains <SUBJECTCONTAINS>           SubjectContains
        --received-after <RECEIVEDAFTER>               ReceivedAfter
        --received-before <RECEIVEDBEFORE>             ReceivedBefore
        --has-attachment <HASATTACHMENT>               HasAttachment
        --flagged-only <FLAGGEDONLY>                   FlaggedOnly
        --cursor <CURSOR>                              Cursor
        --query <QUERY>                                (required for: search)
        --search-mode <SEARCHMODE>                     SearchMode
        --response <RESPONSE>                          Response
        --send-response <SENDRESPONSE>                 SendResponse
        --response-text <RESPONSETEXT>                 ResponseText
        --recipient-to <RECIPIENTTO>                   RecipientTo
        --cc <CC>                                      Cc
        --bcc <BCC>                                    Bcc
        --subject <SUBJECT>                            (required for:
                                                       set-subject)
        --body <BODY>                                  (required for: set-body)
        --display <DISPLAY>                            Display
        --body-format <BODYFORMAT>                     BodyFormat
        --confirm <CONFIRM>                            Confirm
        --operation-id <OPERATIONID>                   OperationId
        --target-folder <TARGETFOLDER>                 (required for: move)
        --file-path <FILEPATH>                         (required for: export)
        --format <FORMAT>                              Format string
        --overwrite <OVERWRITE>                        Overwrite
        --is-read <ISREAD>                             (required for:
                                                       set-read-state)
        --flag-status <FLAGSTATUS>                     FlagStatus
        --due-date <DUEDATE>                           DueDate
        --flag-request <FLAGREQUEST>                   FlagRequest
        --categories <CATEGORIES>                      Categories
        --include-detail <INCLUDEDETAIL>               IncludeDetail
        --upcoming-only <UPCOMINGONLY>                 Keep only reminders that
                                                       have not yet fallen due.
                                                       On by default, because
                                                       most reminders on a
                                                       long-lived mailbox are
                                                       years overdue and
                                                       including them buries the
                                                       ones that matter
    -o, --output <PATH>                                Write output to file
                                                       instead of stdout. For
                                                       image results, decodes
                                                       and saves as binary file

```

## rule

```
DESCRIPTION:
Outlook inbox rule operations

USAGE:
    outlookcli rule <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                           Prints help information
        --include-detail <INCLUDEDETAIL>                 Gather each rule's
                                                         conditions, actions,
                                                         subject terms, sender
                                                         addresses and move-to
                                                         destination. Off by
                                                         default: Outlook's
                                                         condition and action
                                                         collections have a
                                                         fixed length covering
                                                         every clause it
                                                         supports, so detail
                                                         means walking roughly
                                                         59 slots per rule
        --store-id <STOREID>                             The mailbox to read,
                                                         from folder
                                                         list-stores. Defaults
                                                         to the profile's
                                                         default delivery store
        --name <NAME>                                    The rule's name, as it
                                                         will appear in Outlook.
                                                         Must not already be in
                                                         use in this store.
                                                         (required for: create,
                                                         update, set-enabled,
                                                         delete)
        --from-address <FROMADDRESS>                     Match when the sender's
                                                         SMTP address contains
                                                         this. A substring match
                                                         on the address itself,
                                                         so no address-book
                                                         lookup is involved
        --subject-contains <SUBJECTCONTAINS>             Match when the subject
                                                         contains this
        --move-to-folder <MOVETOFOLDER>                  Move matching mail to
                                                         this folder - a default
                                                         folder role such as
                                                         'inbox' or a full
                                                         folder path. The folder
                                                         must already exist
        --assign-categories <ASSIGNCATEGORIES>           Stamp matching mail
                                                         with these categories,
                                                         comma-separated. Use
                                                         mail list-categories to
                                                         discover which names
                                                         exist
        --delete-message <DELETEMESSAGE>                 Move matching mail to
                                                         Deleted Items. Never a
                                                         permanent delete.
                                                         Outlook stores this as
                                                         a move plus
                                                         stop-processing rather
                                                         than as a delete
                                                         action, so list reports
                                                         it as moveToFolder.
                                                         Cannot be combined with
                                                         , because a rule has
                                                         only one move
                                                         destination
        --stop-processing-rules <STOPPROCESSINGRULES>    Stop evaluating later
                                                         rules once this one
                                                         matches
        --enabled <ENABLED>                              Whether the rule acts
                                                         on mail immediately.
                                                         True by default
        --new-name <NEWNAME>                             Rename the rule. Must
                                                         not collide with
                                                         another rule in the
                                                         store
    -o, --output <PATH>                                  Write output to file
                                                         instead of stdout. For
                                                         image results, decodes
                                                         and saves as binary
                                                         file

```

## task

```
DESCRIPTION:
Outlook task operations

USAGE:
    outlookcli task <ACTION> [OPTIONS]

ARGUMENTS:
    <ACTION>    The action to perform

OPTIONS:
    -h, --help                                         Prints help information
        --folder <FOLDER>                              Folder
        --max-count <MAXCOUNT>                         MaxCount
        --include-completed <INCLUDECOMPLETED>         IncludeCompleted
        --include-body-preview <INCLUDEBODYPREVIEW>    IncludeBodyPreview
        --entry-id <ENTRYID>                           EntryId
        --store-id <STOREID>                           StoreId
        --use-active-task <USEACTIVETASK>              UseActiveTask
        --subject <SUBJECT>                            (required for: create)
        --due-date <DUEDATE>                           DueDate
        --start-date <STARTDATE>                       StartDate
        --status <STATUS>                              Status
        --percent-complete <PERCENTCOMPLETE>           PercentComplete
        --importance <IMPORTANCE>                      Importance
        --categories <CATEGORIES>                      Categories
        --body <BODY>                                  Body
        --display <DISPLAY>                            Display
    -o, --output <PATH>                                Write output to file
                                                       instead of stdout. For
                                                       image results, decodes
                                                       and saves as binary file

```
