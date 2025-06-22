let
    // Charger tous les fichiers du dossier
    Source = Folder.Files("P:\RE\COMMUN\SUIVI_PROD_PS\data\2025\juillet"),
    #"Fichiers masqués filtrés1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Appeler une fonction personnalisée1" = Table.AddColumn(#"Fichiers masqués filtrés1", "Transformer le fichier (2)", each #"Transformer le fichier (2)"([Content])),
    #"Colonnes renommées1" = Table.RenameColumns(#"Appeler une fonction personnalisée1", {"Name", "Source.Name"}),
    #"Autres colonnes supprimées1" = Table.SelectColumns(#"Colonnes renommées1", {"Source.Name", "Transformer le fichier (2)"}),
    #"Colonne de tables développée1" = Table.ExpandTableColumn(#"Autres colonnes supprimées1", "Transformer le fichier (2)", Table.ColumnNames(#"Transformer le fichier (2)"(#"Exemple de fichier (2)"))),
    #"Type modifié" = Table.TransformColumnTypes(#"Colonne de tables développée1",{{"Source.Name", type text}, {"FLUX ", type any}, {"Column2", type any}, {"Column3", type any}, {"Column4", type text}}),
    
    // Fonction pour extraire une valeur spécifique basée sur le nom de ligne pour une table donnée
    GetValueFromTable = (table as table, searchText as text, columnName as text) =>
        try
            let
                FilteredRows = Table.SelectRows(table, each 
                    [#"FLUX "] <> null and 
                    Text.Contains(Text.Upper(Text.From([#"FLUX "])), Text.Upper(searchText))
                ),
                FirstRow = Table.FirstN(FilteredRows, 1),
                Value = if Table.RowCount(FirstRow) > 0 then
                    Record.Field(FirstRow{0}, columnName)
                else
                    null
            in
                if Value = null or Value = "" then null else Value
        otherwise null,
    
    // Fonction pour extraire la date d'une table
    GetDateFromTable = (table as table) =>
        try
            let
                DateRow = Table.SelectRows(table, each 
                    try 
                        Date.From([#"FLUX "]) <> null
                    otherwise 
                        false
                ),
                ExtractedDate = if Table.RowCount(DateRow) > 0 then
                    Date.From(DateRow{0}[#"FLUX "])
                else
                    null
            in
                ExtractedDate
        otherwise null,
    
    // Grouper par fichier source et transformer chaque groupe
    GroupedByFile = Table.Group(#"Type modifié", {"Source.Name"}, {
        {"FileData", each _, type table}
    }),
    
    // Transformer chaque fichier
    TransformEachFile = Table.AddColumn(GroupedByFile, "TransformedData", each
        let
            CurrentTable = [FileData],
            
            // Extraire la date
            ExtractedDate = GetDateFromTable(CurrentTable),
            
            // Extraire les valeurs selon le mapping
            ETP_OuvertureIA = GetValueFromTable(CurrentTable, "OUVERTURE IA", "Column2"),
            Prod_OuvertureIA = GetValueFromTable(CurrentTable, "OUVERTURE IA", "Column3"),
            
            ETP_OuvertureManuelle = GetValueFromTable(CurrentTable, "OUVERTURE MANUELLE", "Column2"),
            Prod_OuvertureManuelle = GetValueFromTable(CurrentTable, "OUVERTURE MANUELLE", "Column3"),
            
            ETP_TicketsRetour = GetValueFromTable(CurrentTable, "TICKETS RETOUR RELANCE", "Column2"),
            Prod_TicketsRetour = GetValueFromTable(CurrentTable, "TICKETS RETOUR RELANCE", "Column3"),
            Comment_TicketsRetour = GetValueFromTable(CurrentTable, "TICKETS RETOUR RELANCE", "Column4"),
            
            ETP_UanCourrierSucc = GetValueFromTable(CurrentTable, "UAN COURRIER SUCCESSION", "Column2"),
            Prod_UanCourrierSucc = GetValueFromTable(CurrentTable, "UAN COURRIER SUCCESSION", "Column3"),
            Comment_UanCourrierSucc = GetValueFromTable(CurrentTable, "UAN COURRIER SUCCESSION", "Column4"),
            
            ETP_Transferts = GetValueFromTable(CurrentTable, "TRANSFERTS", "Column2"),
            Prod_Transferts = GetValueFromTable(CurrentTable, "TRANSFERTS", "Column3"),
            Comment_Transferts = GetValueFromTable(CurrentTable, "TRANSFERTS", "Column4"),
            
            Prod_ControleTransfert = GetValueFromTable(CurrentTable, "CONTROLE TRANSFERT", "Column3"),
            Comment_ControleTransfert = GetValueFromTable(CurrentTable, "CONTROLE TRANSFERT", "Column4"),
            
            ETP_BalSuccession = GetValueFromTable(CurrentTable, "BAL SUCCESSION", "Column2"),
            Prod_BalSuccession = GetValueFromTable(CurrentTable, "BAL SUCCESSION", "Column3"),
            Comment_BalSuccession = GetValueFromTable(CurrentTable, "BAL SUCCESSION", "Column4"),
            
            ETP_UanCourrierRef = GetValueFromTable(CurrentTable, "UAN COURRIER REFERENTIEL", "Column2"),
            Prod_UanCourrierRef = GetValueFromTable(CurrentTable, "UAN COURRIER REFERENTIEL", "Column3"),
            Comment_UanCourrierRef = GetValueFromTable(CurrentTable, "UAN COURRIER REFERENTIEL", "Column4"),
            
            ETP_RefSuspens = GetValueFromTable(CurrentTable, "REF: BAL mails traités", "Column2"),
            Prod_RefSuspens = GetValueFromTable(CurrentTable, "REF: BAL mails traités", "Column3"),
            Comment_RefSuspens = GetValueFromTable(CurrentTable, "REF: BAL mails traités", "Column4"),
            
            ETP_BalRef = GetValueFromTable(CurrentTable, "BAL REFERENTIEL: tri", "Column2"),
            Prod_BalRef = GetValueFromTable(CurrentTable, "BAL REFERENTIEL: tri", "Column3"),
            Comment_BalRef = GetValueFromTable(CurrentTable, "BAL REFERENTIEL: tri", "Column4"),
            
            ETP_FormationRef = GetValueFromTable(CurrentTable, "FORMATION REFERENTIEL", "Column2"),
            Comment_FormationRef = GetValueFromTable(CurrentTable, "FORMATION REFERENTIEL", "Column4"),
            
            ETP_PlateformeTel = GetValueFromTable(CurrentTable, "PLATEFORME TEL", "Column2"),
            Prod_PlateformeTel = GetValueFromTable(CurrentTable, "PLATEFORME TEL", "Column3"),
            Comment_PlateformeTel = GetValueFromTable(CurrentTable, "PLATEFORME TEL", "Column4"),
            
            ETP_Management = GetValueFromTable(CurrentTable, "MANAGEMENT", "Column2"),
            Comment_Management = GetValueFromTable(CurrentTable, "MANAGEMENT", "Column4"),
            
            ETP_Absences = GetValueFromTable(CurrentTable, "ABSENCES", "Column2"),
            Comment_Absences = GetValueFromTable(CurrentTable, "ABSENCES", "Column4"),
            
            ETP_Total = GetValueFromTable(CurrentTable, "ETP", "Column2"),
            
            // Calculer les moyennes de productivité
            MoyProd_Ouvertures = if ETP_OuvertureManuelle <> null and ETP_OuvertureIA <> null and 
                                    Prod_OuvertureManuelle <> null and Prod_OuvertureIA <> null then
                try
                    (Number.From(Prod_OuvertureManuelle) + Number.From(Prod_OuvertureIA)) / 
                    (Number.From(ETP_OuvertureManuelle) + Number.From(ETP_OuvertureIA))
                otherwise null
            else null,
            
            MoyProd_TicketsRetour = if ETP_TicketsRetour <> null and Prod_TicketsRetour <> null then
                try Number.From(Prod_TicketsRetour) / Number.From(ETP_TicketsRetour) otherwise null
            else null,
            
            MoyProd_Transferts = if ETP_Transferts <> null and Prod_Transferts <> null then
                try Number.From(Prod_Transferts) / Number.From(ETP_Transferts) otherwise null
            else null,
            
            MoyProd_ControleTransferts = if Prod_ControleTransfert <> null then
                try Number.From(Prod_ControleTransfert) otherwise null
            else null,
            
            MoyProd_BalSucc = if ETP_BalSuccession <> null and Prod_BalSuccession <> null then
                try Number.From(Prod_BalSuccession) / Number.From(ETP_BalSuccession) otherwise null
            else null,
            
            MoyProd_UanSucc = if ETP_UanCourrierSucc <> null and Prod_UanCourrierSucc <> null then
                try Number.From(Prod_UanCourrierSucc) / Number.From(ETP_UanCourrierSucc) otherwise null
            else null,
            
            MoyProd_UanBalRef = if ETP_RefSuspens <> null and Prod_RefSuspens <> null and 
                                  ETP_BalRef <> null and Prod_BalRef <> null then
                try
                    (Number.From(Prod_RefSuspens) + Number.From(Prod_BalRef)) / 
                    (Number.From(ETP_RefSuspens) + Number.From(ETP_BalRef))
                otherwise null
            else null,
            
            // Créer l'enregistrement de résultat
            ResultRecord = [
                date = ExtractedDate,
                ETP = if ETP_OuvertureManuelle <> null and ETP_OuvertureIA <> null then 
                    try Number.From(ETP_OuvertureManuelle) + Number.From(ETP_OuvertureIA) otherwise null else null,
                #"OUVERTURES MANUELLES" = try Number.From(Prod_OuvertureManuelle) otherwise null,
                COMMENTAIRES = null,
                #"DOSSIERS OUVERTURES IA" = try Number.From(Prod_OuvertureIA) otherwise null,
                COMMENTAIRES2 = null,
                #"MOYENNE PROD OUVERTURES" = MoyProd_Ouvertures,
                ETP2 = try Number.From(ETP_TicketsRetour) otherwise null,
                #"TICKETS RETOUR RELANCE PM" = try Number.From(Prod_TicketsRetour) otherwise null,
                COMMENTAIRES3 = Comment_TicketsRetour,
                #"MOYENNE PROD TICKET RETOUR RELANCE PM" = MoyProd_TicketsRetour,
                ETP3 = try Number.From(ETP_Transferts) otherwise null,
                TRANSFERTS = try Number.From(Prod_Transferts) otherwise null,
                COMMENTAIRES4 = Comment_Transferts,
                #"MOYENNE PROD TRANSFERTS" = MoyProd_Transferts,
                ETP4 = null,
                #"CONTRÔLE TRANSFERTS" = try Number.From(Prod_ControleTransfert) otherwise null,
                COMMENTAIRES5 = Comment_ControleTransfert,
                #"MOYENNE PROD CONTRÔLES TRANSFERTS" = MoyProd_ControleTransferts,
                ETP5 = null,
                #"RELANCES PM SUCC" = null,
                COMMENTAIRES6 = null,
                #"MOYENNE PROD RELANCES PM" = null,
                ETP6 = try Number.From(ETP_BalSuccession) otherwise null,
                #"BAL SUCC" = try Number.From(Prod_BalSuccession) otherwise null,
                COMMENTAIRES7 = Comment_BalSuccession,
                #"MOYENNE PROD BAL SUCC" = MoyProd_BalSucc,
                ETP7 = try Number.From(ETP_UanCourrierSucc) otherwise null,
                #"TICKETS UAN COURRIERS SUCC" = try Number.From(Prod_UanCourrierSucc) otherwise null,
                COMMENTAIRES8 = Comment_UanCourrierSucc,
                #"MOYENNE PROD UAN SUCC" = MoyProd_UanSucc,
                ETP8 = try Number.From(ETP_FormationRef) otherwise null,
                #"REFERENTIEL (formation/ contrôle des taches)" = null,
                COMMENTAIRES9 = Comment_FormationRef,
                ETP9 = try Number.From(ETP_BalRef) otherwise null,
                #"BAL REFERENTIEL" = try Number.From(Prod_BalRef) otherwise null,
                COMMENTAIRES10 = Comment_BalRef,
                ETP10 = try Number.From(ETP_UanCourrierRef) otherwise null,
                #"UAN REFERENTIEL" = try Number.From(Prod_UanCourrierRef) otherwise null,
                COMMENTAIRES11 = Comment_UanCourrierRef,
                ETP11 = try Number.From(ETP_RefSuspens) otherwise null,
                #"REF SUSPENS/REJETS" = try Number.From(Prod_RefSuspens) otherwise null,
                COMMENTAIRES12 = Comment_RefSuspens,
                #"MOYENNE PROD UAN BAL REF" = MoyProd_UanBalRef,
                ETP12 = try Number.From(ETP_PlateformeTel) otherwise null,
                #"PT TEL" = try Number.From(Prod_PlateformeTel) otherwise null,
                COMMENTAIRES13 = Comment_PlateformeTel,
                #"ETP MANAGEMENT" = try Number.From(ETP_Management) otherwise null,
                COMMENTAIRES14 = Comment_Management,
                #"ETP ABSENCE" = try Number.From(ETP_Absences) otherwise null,
                COMMENTAIRES15 = Comment_Absences,
                #"TOTAL ETP" = try Number.From(ETP_Total) otherwise null
            ]
        in
            ResultRecord
    ),
    
    // Développer les données transformées
    ExpandTransformed = Table.ExpandRecordColumn(TransformEachFile, "TransformedData", 
        {"date", "ETP", "OUVERTURES MANUELLES", "COMMENTAIRES", "DOSSIERS OUVERTURES IA", "COMMENTAIRES2", 
         "MOYENNE PROD OUVERTURES", "ETP2", "TICKETS RETOUR RELANCE PM", "COMMENTAIRES3", 
         "MOYENNE PROD TICKET RETOUR RELANCE PM", "ETP3", "TRANSFERTS", "COMMENTAIRES4", 
         "MOYENNE PROD TRANSFERTS", "ETP4", "CONTRÔLE TRANSFERTS", "COMMENTAIRES5", 
         "MOYENNE PROD CONTRÔLES TRANSFERTS", "ETP5", "RELANCES PM SUCC", "COMMENTAIRES6", 
         "MOYENNE PROD RELANCES PM", "ETP6", "BAL SUCC", "COMMENTAIRES7", "MOYENNE PROD BAL SUCC", 
         "ETP7", "TICKETS UAN COURRIERS SUCC", "COMMENTAIRES8", "MOYENNE PROD UAN SUCC", 
         "ETP8", "REFERENTIEL (formation/ contrôle des taches)", "COMMENTAIRES9", "ETP9", 
         "BAL REFERENTIEL", "COMMENTAIRES10", "ETP10", "UAN REFERENTIEL", "COMMENTAIRES11", 
         "ETP11", "REF SUSPENS/REJETS", "COMMENTAIRES12", "MOYENNE PROD UAN BAL REF", 
         "ETP12", "PT TEL", "COMMENTAIRES13", "ETP MANAGEMENT", "COMMENTAIRES14", 
         "ETP ABSENCE", "COMMENTAIRES15", "TOTAL ETP"}),
    
    // Supprimer les colonnes intermédiaires et ne garder que les résultats
    SelectFinalColumns = Table.SelectColumns(ExpandTransformed, 
        {"date", "ETP", "OUVERTURES MANUELLES", "COMMENTAIRES", "DOSSIERS OUVERTURES IA", "COMMENTAIRES2", 
         "MOYENNE PROD OUVERTURES", "ETP2", "TICKETS RETOUR RELANCE PM", "COMMENTAIRES3", 
         "MOYENNE PROD TICKET RETOUR RELANCE PM", "ETP3", "TRANSFERTS", "COMMENTAIRES4", 
         "MOYENNE PROD TRANSFERTS", "ETP4", "CONTRÔLE TRANSFERTS", "COMMENTAIRES5", 
         "MOYENNE PROD CONTRÔLES TRANSFERTS", "ETP5", "RELANCES PM SUCC", "COMMENTAIRES6", 
         "MOYENNE PROD RELANCES PM", "ETP6", "BAL SUCC", "COMMENTAIRES7", "MOYENNE PROD BAL SUCC", 
         "ETP7", "TICKETS UAN COURRIERS SUCC", "COMMENTAIRES8", "MOYENNE PROD UAN SUCC", 
         "ETP8", "REFERENTIEL (formation/ contrôle des taches)", "COMMENTAIRES9", "ETP9", 
         "BAL REFERENTIEL", "COMMENTAIRES10", "ETP10", "UAN REFERENTIEL", "COMMENTAIRES11", 
         "ETP11", "REF SUSPENS/REJETS", "COMMENTAIRES12", "MOYENNE PROD UAN BAL REF", 
         "ETP12", "PT TEL", "COMMENTAIRES13", "ETP MANAGEMENT", "COMMENTAIRES14", 
         "ETP ABSENCE", "COMMENTAIRES15", "TOTAL ETP"}),
    
    // Trier par date
    SortedByDate = Table.Sort(SelectFinalColumns, {{"date", Order.Ascending}}),
    
    // Définir les types de colonnes finaux (format DATE pour la colonne date)
    FinalTypes = Table.TransformColumnTypes(SortedByDate, {
        {"date", type nullable date},
        {"ETP", type nullable number},
        {"OUVERTURES MANUELLES", type nullable number},
        {"COMMENTAIRES", type nullable text},
        {"DOSSIERS OUVERTURES IA", type nullable number},
        {"COMMENTAIRES2", type nullable text},
        {"MOYENNE PROD OUVERTURES", type nullable number},
        {"ETP2", type nullable number},
        {"TICKETS RETOUR RELANCE PM", type nullable number},
        {"COMMENTAIRES3", type nullable text},
        {"MOYENNE PROD TICKET RETOUR RELANCE PM", type nullable number},
        {"ETP3", type nullable number},
        {"TRANSFERTS", type nullable number},
        {"COMMENTAIRES4", type nullable text},
        {"MOYENNE PROD TRANSFERTS", type nullable number},
        {"ETP4", type nullable number},
        {"CONTRÔLE TRANSFERTS", type nullable number},
        {"COMMENTAIRES5", type nullable text},
        {"MOYENNE PROD CONTRÔLES TRANSFERTS", type nullable number},
        {"ETP5", type nullable number},
        {"RELANCES PM SUCC", type nullable number},
        {"COMMENTAIRES6", type nullable text},
        {"MOYENNE PROD RELANCES PM", type nullable number},
        {"ETP6", type nullable number},
        {"BAL SUCC", type nullable number},
        {"COMMENTAIRES7", type nullable text},
        {"MOYENNE PROD BAL SUCC", type nullable number},
        {"ETP7", type nullable number},
        {"TICKETS UAN COURRIERS SUCC", type nullable number},
        {"COMMENTAIRES8", type nullable text},
        {"MOYENNE PROD UAN SUCC", type nullable number},
        {"ETP8", type nullable number},
        {"REFERENTIEL (formation/ contrôle des taches)", type nullable number},
        {"COMMENTAIRES9", type nullable text},
        {"ETP9", type nullable number},
        {"BAL REFERENTIEL", type nullable number},
        {"COMMENTAIRES10", type nullable text},
        {"ETP10", type nullable number},
        {"UAN REFERENTIEL", type nullable number},
        {"COMMENTAIRES11", type nullable text},
        {"ETP11", type nullable number},
        {"REF SUSPENS/REJETS", type nullable number},
        {"COMMENTAIRES12", type nullable text},
        {"MOYENNE PROD UAN BAL REF", type nullable number},
        {"ETP12", type nullable number},
        {"PT TEL", type nullable number},
        {"COMMENTAIRES13", type nullable text},
        {"ETP MANAGEMENT", type nullable number},
        {"COMMENTAIRES14", type nullable text},
        {"ETP ABSENCE", type nullable number},
        {"COMMENTAIRES15", type nullable text},
        {"TOTAL ETP", type nullable number}
    })

in
    FinalTypes