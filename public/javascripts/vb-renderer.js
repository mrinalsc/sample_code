Option Compare Database
Dim abort_import As Boolean
Global LibraryXML As String
Global GroupsXML As String
Global requiredEffectivity As String
Global import_options As String
Dim ManifestXMLFile As String
Global CoreXMLFile As String
Global PR As String
'Global vl As String
Global CoreXMLID As Long
Global newBMCCode As String
Global CoreXMLBranch As String
Global MYID As String
Global BMCID As String
Global FASpanID As String
Global ProgrammeID As String
Global RuleID As String
Global SUID As String

Global Import_Mode As String
Global Debug_Mode As Boolean
Dim eventlists As Variant
Global e As Integer
Global xmlEvents As Scripting.Dictionary

Dim FA_market As Scripting.Dictionary
Dim FA_bmc As Scripting.Dictionary
Dim FA_feature As Scripting.Dictionary
Dim XML_Events As Scripting.Dictionary

Dim XMLDesc As String

Dim isSubModel As Boolean

Dim altCodes As Scripting.Dictionary

Global impBasics As Boolean
Global impProgrammes As Boolean
Global impEvents As Boolean
Global impModelYears As Boolean
Global impOK2UL As Boolean
Global impBM As Boolean
Global impMA As Boolean
Global impFA As Boolean
Global impRules As Boolean
Global impSU As Boolean
Global impSUOnly As Boolean
Global impTableRules As Boolean
Dim impExists As Boolean
Global rMFDDimension As Scripting.Dictionary



Sub LoadCoreXMLFile(ManifestXMLFile As String, CoreXMLFile As String, Optional requiredEffectivity As String, Optional inXMLDesc As String)

    Dim tempFeature As String
    Dim vl As String
    
    XMLDesc = inXMLDesc
    
    isSubModel = True
    impMA = True
    impExists = False
     
    Import_Mode = "RST"
    'Import_Mode = "DOCMD"
    Debug_Mode = False
    
    Debug.Print Import_Mode
    Debug.Print "Started: " & Now

    DoCmd.SetWarnings True
    
    Set dbs = CurrentDb
    get_feature_library
    

    Set XML_Events = New Scripting.Dictionary

    Set XDocMani = CreateObject("Msxml2.DOMDocument.6.0")
    XDocMani.async = False: XDocMani.validateOnParse = False
    XDocMani.Load (ManifestXMLFile)
    
    Set manilists = XDocMani.DocumentElement
    
    For Each maniDetails In manilists.ChildNodes
        If maniDetails.BaseName = "Metadata" Then
            For Each manidetail In maniDetails.ChildNodes
                If manidetail.Attributes.getnameditem("Key").Text = "Branch" Then
                    CoreXMLBranch = manidetail.Attributes.getnameditem("Value").Text
                ElseIf manidetail.Attributes.getnameditem("Key").Text = "Revision" Then
                    CoreXMLID = manidetail.Attributes.getnameditem("Value").Text
                End If
            Next
        End If
        
        If maniDetails.BaseName = "ImportDetails" Then
            For Each addinfo In maniDetails.ChildNodes
                If addinfo.BaseName = "AdditionalInformation" Then
                    For Each AddInfoDetail In addinfo.ChildNodes
                        If AddInfoDetail.Attributes.getnameditem("Key").Text = "Branch" Then
                            CoreXMLBranch = AddInfoDetail.Attributes.getnameditem("Value").Text
                        ElseIf AddInfoDetail.Attributes.getnameditem("Key").Text = "Revision" Then
                            CoreXMLID = AddInfoDetail.Attributes.getnameditem("Value").Text
                        End If
                    Next
                End If
            Next
        End If
        If maniDetails.BaseName = "Files" Then
            For Each fileinfo In maniDetails.ChildNodes
                If fileinfo.BaseName = "AdditioFGnalInformation" Then
                    For Each AddInfoDetail In addinfo.ChildNodes
                        If AddInfoDetail.Attributes.getnameditem("Key").Text = "Branch" Then
                            CoreXMLBranch = AddInfoDetail.Attributes.getnameditem("Value").Text
                        ElseIf AddInfoDetail.Attributes.getnameditem("Key").Text = "Revision" Then
                            CoreXMLID = AddInfoDetail.Attributes.getnameditem("Value").Text
                        End If
                    Next
                End If
            Next
        End If
    Next
    PR = vl
    
        
    Set XDoc = CreateObject("Msxml2.DOMDocument.6.0")
    XDoc.async = False: XDoc.validateOnParse = True

    Debug.Print "Started Loading XML: " & Now, CoreXMLFile
    XDoc.Load (CoreXMLFile)
    Debug.Print "Finished Loading XML: " & Now
    

    Set Lists = XDoc.DocumentElement
    
    For Each productranges In Lists.ChildNodes
        
        If productranges.BaseName = "ProductRanges" Or productranges.BaseName = "ProductModels" Then
            For Each productrange In productranges.ChildNodes
                For Each Branch In productrange.ChildNodes
                    'If impExists = False Then
                    
                        If Branch.BaseName = "Basics" Then
                            temp1 = productrange.Attributes.getnameditem("Code").Text
                            temp2 = Branch.Attributes.getnameditem("Description").Text
                            temp3 = Branch.Attributes.getnameditem("Brand").Text
                            temp4 = Branch.Attributes.getnameditem("Feature").Text
                            vl = productrange.Attributes.getnameditem("Code").Text
                            
                            add_CoreXML_Details CoreXMLFile, vl, requiredEffectivity

                            If abort_import = True Then Exit Sub
                            
                            add_basics temp1, temp2, temp3, temp4
                        End If
                        If Branch.BaseName = "Programmes" Then
                            add_programme CoreXMLID, "START", "Start Programme"
                            For Each Programme In Branch.ChildNodes
                                add_programme CoreXMLID, Programme.Attributes.getnameditem("Code").Text, Programme.Attributes.getnameditem("Description").Text
                            Next
                            add_programme CoreXMLID, "END", "END Programme"
                            isSubModel = False
                        End If
                        If Branch.BaseName = "Events" Then
                            If isSubModel = True Then
                                add_programme CoreXMLID, "SUB", "Sub Model Programme"
                            End If
                            
                            
                            ReDim eventlists(1 To Branch.ChildNodes.Length + 2, 1 To 5)
                            e = 1
                            add_events CoreXMLID, vl & "_START", "Default Start", "2000-01-01 00:00:00", "DefaultStart", "START", "2000-01-01 00:00:00"
                            For Each events In Branch.ChildNodes
                                If isSubModel = False Then
                                    add_events CoreXMLID, events.Attributes.getnameditem("Code").Text, events.Attributes.getnameditem("Description").Text, events.Attributes.getnameditem("Date").Text, events.Attributes.getnameditem("EventType").Text, events.Attributes.getnameditem("Programme").Text
                                Else
                                    add_events CoreXMLID, events.Attributes.getnameditem("Code").Text, events.Attributes.getnameditem("Description").Text, events.Attributes.getnameditem("Date").Text, events.Attributes.getnameditem("EventType").Text, "SUB"
                                End If
                            Next
                            add_events CoreXMLID, vl & "_END", "Default END", "9999-12-31 00:00:00", "DefaultEnd", "END", "9999-12-31 00:00:00"
                        
                        
                            For job1 = 1 To UBound(eventlists, 1)
                                
                                If eventlists(job1, 2) = "VolumeIn" Then
                                    job1Programme = eventlists(job1, 3)
                                    job1date = eventlists(job1, 4)
                                    
                                    'tempT = Format(Left(job1date, 10), "YYYY-MM-DD") & " " & Format(Mid(job1date, 12, 8), "HH:MM:SS")
                                    eventlists(job1, 5) = DateAdd("n", 120, job1date)
                                    preVolSeconds = 5
                                    rcVolSeconds = 125
                                    For b = 1 To UBound(eventlists, 1)
                                        If eventlists(b, 2) = "PreVolume" And eventlists(b, 3) = job1Programme Then
                                            
                                            eventlists(b, 5) = DateAdd("n", preVolSeconds, job1date)
                                            preVolSeconds = preVolSeconds + 1
                                        End If
                                        If eventlists(b, 2) = "Running" And eventlists(b, 3) = job1Programme Then
                                            
                                            eventlists(b, 5) = DateAdd("n", rcVolSeconds, job1date)
                                            rcVolSeconds = rcVolSeconds + 1
                                        End If
                                    Next b
                                    
                                End If
                            
                                Next job1
restart_effectivity:
                                
                                e = 0
                                For A = 1 To UBound(eventlists, 1)
                                    
                                    Dim myEvent As String
                                    myEvent = eventlists(A, 1)
                                    If eventlists(A, 2) = "VolumeOut" Then
                                        newDate = eventlists(A, 5)
                                    Else
                                        newDate = eventlists(A, 5)
                                    End If
                                    
                                    If Left(eventlists(A, 2), 7) <> "Default" Then
                                        update_event_sequence CoreXMLID, myEvent, newDate
                                    End If
                                Next A
                                
                                create_xmlEventsevents_list
                                
                                If requiredEffectivityDate = "" Then
                                    requiredEffectivityDate = Format(xmlEvents(requiredEffectivity), "YYYY/MM/DD HH:MM:SS")
                                End If
    
                        End If
                    'End If
                    If impSUOnly = False Then
                        If Branch.BaseName = "Associations" Then
                            For Each Associations In Branch.ChildNodes
                                If Associations.BaseName = "FeatureAssociations" Then
                                    For Each Association In Associations.ChildNodes
                                        For Each Effectivities In Association.ChildNodes
                                            For Each effectivity In Effectivities.ChildNodes
                                                Dim spanstart As String
                                                Dim spanend As String
                                                spanstart = effectivity.Attributes.getnameditem("StartEvent").Text
                                                spanend = effectivity.Attributes.getnameditem("EndEvent").Text
                                                spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                                spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                                
                                                If requiredEffectivityDate = "" Then
                                                    add_FeatureAssociations CoreXMLID, Association.Attributes.getnameditem("Code").Text, Association.Attributes.getnameditem("FeatureFamily").Text, Association.Attributes.getnameditem("Priority").Text, effectivity.Attributes.getnameditem("StartEvent").Text, effectivity.Attributes.getnameditem("EndEvent").Text
                                                Else
                                                    If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                        add_FeatureAssociations CoreXMLID, Association.Attributes.getnameditem("Code").Text, Association.Attributes.getnameditem("FeatureFamily").Text, Association.Attributes.getnameditem("Priority").Text, effectivity.Attributes.getnameditem("StartEvent").Text, effectivity.Attributes.getnameditem("EndEvent").Text
                                                    End If
                                                End If
                                            Next
                                        Next
                                    Next
                                End If
                                If Associations.BaseName = "FeatureFamilyAssociations" Then
                                    For Each FamilyAssociation In Associations.ChildNodes
                                        'Debug.Print FamilyAssociation.Attributes.getnameditem("Code").Text
                                        'If FamilyAssociation.BaseName = "Features" Then
                                            For Each FeatAssociation In FamilyAssociation.ChildNodes
                                                If FeatAssociation.BaseName = "Features" Then
                                                    For Each Feature In FeatAssociation.ChildNodes
                                                        'Debug.Print Feature.Attributes.getnameditem("Code").Text
                                                        For Each Effectivities In Feature.ChildNodes
                                                            For Each effectivity In Effectivities.ChildNodes
                                                                spanstart = effectivity.Attributes.getnameditem("StartEvent").Text
                                                                spanend = effectivity.Attributes.getnameditem("EndEvent").Text
                                                                spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                                                spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                                                
                                                                If requiredEffectivityDate = "" Then
                                                                    add_FeatureAssociations CoreXMLID, Feature.Attributes.getnameditem("Code").Text, FamilyAssociation.Attributes.getnameditem("Code").Text, 1, effectivity.Attributes.getnameditem("StartEvent").Text, effectivity.Attributes.getnameditem("EndEvent").Text
                                                                Else
                                                                    If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                                        add_FeatureAssociations CoreXMLID, Feature.Attributes.getnameditem("Code").Text, FamilyAssociation.Attributes.getnameditem("Code").Text, 1, effectivity.Attributes.getnameditem("StartEvent").Text, effectivity.Attributes.getnameditem("EndEvent").Text
                                                                    End If
                                                                End If
                                                            Next
                                                        Next
                                                    Next
                                                End If
                                            Next
                                        'End If
                                        
                                    Next
                                End If
                            Next
                        End If
                    End If
                        If Branch.BaseName = "ModelYearMappings" Then
                            For Each ModelYearMappings In Branch.ChildNodes
                                add_ModelYearMappings CoreXMLID, ModelYearMappings.Attributes.getnameditem("Feature").Text, ModelYearMappings.Attributes.getnameditem("FeatureFamily").Text, ModelYearMappings.Attributes.getnameditem("StartEvent").Text, ModelYearMappings.Attributes.getnameditem("EndEvent").Text
                                For Each BuildPhaseMapping In ModelYearMappings.ChildNodes
                                    For Each BuildPhase In BuildPhaseMapping.ChildNodes
                                        
                                        add_BuildPhases CoreXMLID, BuildPhase.Attributes.getnameditem("Feature").Text, BuildPhase.Attributes.getnameditem("FeatureFamily").Text, BuildPhase.Attributes.getnameditem("BuildPhaseType").Text, BuildPhase.Attributes.getnameditem("StartEvent").Text, BuildPhase.Attributes.getnameditem("EndEvent").Text
                                    Next
                                Next
                            Next
                        End If
    '-----------------------------------------------------------------------------------
    '-----------  Brochure Models and Derivatives --------------------------------------
                    If impSUOnly = False Then
                        If Branch.BaseName = "BrochureModels" Then
                            For Each BrochureModels In Branch.ChildNodes
                                
                                For Each BrochureModel In BrochureModels.ChildNodes
                                    If BrochureModel.BaseName = "BrochureModelDefiningFeatures" Then
                                        For Each DefFeat In BrochureModel.ChildNodes
                                            If DefFeat.Attributes.getnameditem("FamilyCode").Text = "AADA" Then
                                                AADA = DefFeat.Attributes.getnameditem("Code").Text
                                            ElseIf DefFeat.Attributes.getnameditem("FamilyCode").Text = "AAGA" Then
                                                AAGA = DefFeat.Attributes.getnameditem("Code").Text
                                            ElseIf DefFeat.Attributes.getnameditem("FamilyCode").Text = "PBHA" Then
                                                PBHA = DefFeat.Attributes.getnameditem("Code").Text
                                            ElseIf DefFeat.Attributes.getnameditem("FamilyCode").Text = "ZZUW" Then
                                                ZZUW = DefFeat.Attributes.getnameditem("Code").Text
                                            End If
                                        Next DefFeat
                                        add_BrochureModels CoreXMLID, BrochureModels.Attributes.getnameditem("Code").Text, BrochureModels.Attributes.getnameditem("Description").Text, BrochureModels.Attributes.getnameditem("Feature").Text, BrochureModels.Attributes.getnameditem("FeatureFamily").Text, AADA, AAGA, PBHA, ZZUW
                                    End If
    
                                    If BrochureModel.BaseName = "Derivatives" Then
                                        For Each Derivative In BrochureModel.ChildNodes
                                            add_Derivatives CoreXMLID, Derivative.Attributes.getnameditem("Code").Text, Derivative.Attributes.getnameditem("FamilyCode").Text
                                        Next
                                    End If
                                Next
                            Next
                        End If
    '-----------------------------------------------------------------------------------
    '-----------  MARKET AVAILABILITY --------------------------------------------------
                        If Branch.BaseName = "MarketAvailability" And impMA = True Then
                            For Each MarketAvailability In Branch.ChildNodes
                                spanstart = MarketAvailability.Attributes.getnameditem("StartEvent").Text
                                spanend = MarketAvailability.Attributes.getnameditem("EndEvent").Text
                                spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                
                                If requiredEffectivityDate = "" Then
                                    For Each BrochureModelDerivatives In MarketAvailability.ChildNodes
                                        For Each BrochureModelDerivative In BrochureModelDerivatives.ChildNodes
                                        add_marketAvailability CoreXMLID, MarketAvailability.Attributes.getnameditem("Market").Text, MarketAvailability.Attributes.getnameditem("IsAvailable").Text, MarketAvailability.Attributes.getnameditem("StartEvent").Text, MarketAvailability.Attributes.getnameditem("EndEvent").Text, BrochureModelDerivative.Attributes.getnameditem("BrochureModel").Text, BrochureModelDerivative.Attributes.getnameditem("Derivative").Text
                                        Next
                                    Next
                                Else
                                    If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                        For Each BrochureModelDerivatives In MarketAvailability.ChildNodes
                                            For Each BrochureModelDerivative In BrochureModelDerivatives.ChildNodes
                                                add_marketAvailability CoreXMLID, MarketAvailability.Attributes.getnameditem("Market").Text, MarketAvailability.Attributes.getnameditem("IsAvailable").Text, MarketAvailability.Attributes.getnameditem("StartEvent").Text, MarketAvailability.Attributes.getnameditem("EndEvent").Text, BrochureModelDerivative.Attributes.getnameditem("BrochureModel").Text, BrochureModelDerivative.Attributes.getnameditem("Derivative").Text
                                            Next
                                        Next
                                    End If
                                End If
                            Next
                        End If
    '-----------------------------------------------------------------------------------
    '-----------  FeatureApplicability -------------------------------------------------
                        If Branch.BaseName = "FeatureApplicability" Then
                            For Each FeatureApplicability In Branch.ChildNodes
                                If FeatureApplicability.BaseName = "IncludedFamilies" Then
                                
                                    For Each IncludedFamilies In FeatureApplicability.ChildNodes
                                        If productranges.BaseName = "ProductRanges" Then
                                        
                                            add_IncludedFamilies CoreXMLID, IncludedFamilies.Attributes.getnameditem("Code").Text, IncludedFamilies.Attributes.getnameditem("StartEvent").Text, vl & "_END"
                                    
                                        Else
                                            
                                            For Each ApplicabilityEffectivities In IncludedFamilies.ChildNodes
                                                If ApplicabilityEffectivities.BaseName = "ApplicabilityEffectivities" Then
                                                    For Each AppEff In ApplicabilityEffectivities.ChildNodes
                                                        add_IncludedFamilies CoreXMLID, IncludedFamilies.Attributes.getnameditem("Code").Text, AppEff.Attributes.getnameditem("StartEvent").Text, AppEff.Attributes.getnameditem("EndEvent").Text
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                ElseIf FeatureApplicability.BaseName = "FeatureApplicabilitySpans" And impFA = True Then
                                    For Each FeatureApplicabilitySpans In FeatureApplicability.ChildNodes
                                        spanstart = FeatureApplicabilitySpans.Attributes.getnameditem("StartEvent").Text
                                        spanend = FeatureApplicabilitySpans.Attributes.getnameditem("EndEvent").Text
                                        spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                        spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                        If requiredEffectivityDate = "" Then
                                            'add_FASpan CoreXMLID, FeatureApplicabilitySpans.Attributes.getnameditem("Family").Text, FeatureApplicabilitySpans.Attributes.getnameditem("StartEvent").Text, FeatureApplicabilitySpans.Attributes.getnameditem("EndEvent").Text
            
                                            For Each FeatureApplicabilitySpan In FeatureApplicabilitySpans.ChildNodes
                                                If FeatureApplicabilitySpan.BaseName = "Markets" Then
                                                    Set FA_market = New Scripting.Dictionary
                                                    For Each Market In FeatureApplicabilitySpan.ChildNodes
                                                        'add_FASpanMarket market.Attributes.getnameditem("Code").Text
                                                        FA_market.Add Market.Attributes.getnameditem("Code").Text, 1
                                                    Next
                                                ElseIf FeatureApplicabilitySpan.BaseName = "BrochureModelDerivatives" Then
                                                    Set FA_bmc = New Scripting.Dictionary
                                                    For Each BrochureModelDerivatives In FeatureApplicabilitySpan.ChildNodes
                                                        'add_FASpanBMD BrochureModelDerivatives.Attributes.getnameditem("BrochureModel").Text, BrochureModelDerivatives.Attributes.getnameditem("Derivative").Text
                                                        FA_bmc.Add BrochureModelDerivatives.Attributes.getnameditem("BrochureModel").Text & "|" & BrochureModelDerivatives.Attributes.getnameditem("Derivative").Text, 1
                                                    Next
                                                ElseIf FeatureApplicabilitySpan.BaseName = "Features" Then
                                                    Set FA_feature = New Scripting.Dictionary
                                                    For Each Features In FeatureApplicabilitySpan.ChildNodes
                                                        'add_FASpanFeature Features.Attributes.getnameditem("Code").Text, Features.Attributes.getnameditem("FamilyCode").Text, Features.Attributes.getnameditem("Availability").Text, Features.Attributes.getnameditem("MarketingValue").Text
                                                        FA_feature.Add Features.Attributes.getnameditem("Code").Text, Features.Attributes.getnameditem("FamilyCode").Text & "|" & Features.Attributes.getnameditem("Availability").Text & "|" & Features.Attributes.getnameditem("MarketingValue").Text
                                                    Next
                                                End If
                                            Next
                                            
                                            combined_FA FeatureApplicabilitySpans.Attributes.getnameditem("StartEvent").Text, FeatureApplicabilitySpans.Attributes.getnameditem("EndEvent").Text
                                        Else
                                            If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                'add_FASpan CoreXMLID, FeatureApplicabilitySpans.Attributes.getnameditem("Family").Text, FeatureApplicabilitySpans.Attributes.getnameditem("StartEvent").Text, FeatureApplicabilitySpans.Attributes.getnameditem("EndEvent").Text
            
                                                For Each FeatureApplicabilitySpan In FeatureApplicabilitySpans.ChildNodes
                                                    If FeatureApplicabilitySpan.BaseName = "Markets" Then
                                                        Set FA_market = New Scripting.Dictionary
                                                        For Each Market In FeatureApplicabilitySpan.ChildNodes
                                                            'add_FASpanMarket market.Attributes.getnameditem("Code").Text
                                                            FA_market.Add Market.Attributes.getnameditem("Code").Text, 1
                                                        Next
                                                    ElseIf FeatureApplicabilitySpan.BaseName = "BrochureModelDerivatives" Then
                                                        Set FA_bmc = New Scripting.Dictionary
                                                        For Each BrochureModelDerivatives In FeatureApplicabilitySpan.ChildNodes
                                                            'add_FASpanBMD BrochureModelDerivatives.Attributes.getnameditem("BrochureModel").Text, BrochureModelDerivatives.Attributes.getnameditem("Derivative").Text
                                                            FA_bmc.Add BrochureModelDerivatives.Attributes.getnameditem("BrochureModel").Text & "|" & BrochureModelDerivatives.Attributes.getnameditem("Derivative").Text, 1
                                                        Next
                                                    ElseIf FeatureApplicabilitySpan.BaseName = "Features" Then
                                                        Set FA_feature = New Scripting.Dictionary
                                                        For Each Features In FeatureApplicabilitySpan.ChildNodes
                                                            'add_FASpanFeature Features.Attributes.getnameditem("Code").Text, Features.Attributes.getnameditem("FamilyCode").Text, Features.Attributes.getnameditem("Availability").Text, Features.Attributes.getnameditem("MarketingValue").Text
                                                            FA_feature.Add Features.Attributes.getnameditem("Code").Text, Features.Attributes.getnameditem("FamilyCode").Text & "|" & Features.Attributes.getnameditem("Availability").Text & "|" & Features.Attributes.getnameditem("MarketingValue").Text
                                                        Next
                                                    End If
                                                    
                                                Next
                                                combined_FA FeatureApplicabilitySpans.Attributes.getnameditem("StartEvent").Text, FeatureApplicabilitySpans.Attributes.getnameditem("EndEvent").Text
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If
    '-----------------------------------------------------------------------------------
    '-----------  Rules ----------------------------------------------------------------
                        If Branch.BaseName = "Rules" Then
                            For Each Rules In Branch.ChildNodes
                                If Rules.BaseName = "TextRules" Then
                                    For Each TextRuleDetails In Rules.ChildNodes
                                        
                                        spanstart = TextRuleDetails.Attributes.getnameditem("StartEvent").Text
                                        spanend = TextRuleDetails.Attributes.getnameditem("EndEvent").Text
                                        spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                        spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                        'If requiredEffectivityDate = "" Then
                                        If 1 = 1 Then 'Import all text rules
                                            add_RuleDetails CoreXMLID, TextRuleDetails.Attributes.getnameditem("Code").Text, TextRuleDetails.Attributes.getnameditem("Description").Text, TextRuleDetails.Attributes.getnameditem("Intent").Text, TextRuleDetails.Attributes.getnameditem("StartEvent").Text, TextRuleDetails.Attributes.getnameditem("EndEvent").Text, TextRuleDetails.Attributes.getnameditem("IsEnabled").Text, TextRuleDetails.Attributes.getnameditem("IsLocked").Text, "Text"
                                            For Each TextRule In TextRuleDetails.ChildNodes
                                                If TextRule.BaseName = "Labels" Then
                                                    For Each Label In TextRule.ChildNodes
                                                        add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                        If Label.Attributes.getnameditem("Name").Text = "OXO Rules" Then
                                                            Dim myLabel As String
                                                            myLabel = Label.Attributes.getnameditem("Name").Text
                                                        ElseIf Label.Attributes.getnameditem("Name").Text = "OXO Pack Restrictions" Then
                                                            myLabel = Label.Attributes.getnameditem("Name").Text
                                                        
                                                        End If
                                                    Next
                                                End If
                                                If TextRule.BaseName = "Body" Then
                                                    For Each Body In TextRule.ChildNodes
                                                        tempCleanRuleBody = Split(Body.Text, vbLf, , vbTextCompare)
                                                        tempBody = ""
                                                        For CB = 0 To UBound(tempCleanRuleBody)
                                                            If Left(tempCleanRuleBody(CB), 2) = "//" Then
                                                                tempCleanRuleBody(CB) = Replace(tempCleanRuleBody(CB), "[", "{", , , vbTextCompare)
                                                                tempCleanRuleBody(CB) = Replace(tempCleanRuleBody(CB), "]", "}", , , vbTextCompare)
                                                        
                                                            End If
                                                            tempBody = tempBody & tempCleanRuleBody(CB)
                                                        Next CB
                                                        
                                                        tempRuleBody = Split(tempBody, "[", , vbTextCompare)
                                                        If UBound(tempRuleBody) > 0 Then
                                                            tempFeature = Left(tempRuleBody(1), InStr(1, tempRuleBody(1), "]", vbTextCompare) - 1)
                                                        End If
                                                        add_RuleBody Body.Text, tempFeature, myLabel
                                                        
                                                    Next
                                                End If
                                            Next
                                        Else
                                            If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                add_RuleDetails CoreXMLID, TextRuleDetails.Attributes.getnameditem("Code").Text, TextRuleDetails.Attributes.getnameditem("Description").Text, TextRuleDetails.Attributes.getnameditem("Intent").Text, TextRuleDetails.Attributes.getnameditem("StartEvent").Text, TextRuleDetails.Attributes.getnameditem("EndEvent").Text, TextRuleDetails.Attributes.getnameditem("IsEnabled").Text, TextRuleDetails.Attributes.getnameditem("IsLocked").Text, "Text"
                                                For Each TextRule In TextRuleDetails.ChildNodes
                                                    If TextRule.BaseName = "Labels" Then
                                                        For Each Label In TextRule.ChildNodes
                                                            add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                            If Label.Attributes.getnameditem("Name").Text = "OXO Rules" Then
                                                            
                                                                myLabel = Label.Attributes.getnameditem("Name").Text
                                                            ElseIf Label.Attributes.getnameditem("Name").Text = "OXO Pack Restrictions" Then
                                                                myLabel = Label.Attributes.getnameditem("Name").Text
                                                            
                                                            End If
                                                        Next
                                                    End If
                                                    If TextRule.BaseName = "Body" Then
                                                        For Each Body In TextRule.ChildNodes
                                                            tempCleanRuleBody = Split(Body.Text, vbLf, , vbTextCompare)
                                                            tempBody = ""
                                                            For CB = 0 To UBound(tempCleanRuleBody)
                                                                If Left(tempCleanRuleBody(CB), 2) = "//" Then
                                                                    tempCleanRuleBody(CB) = Replace(tempCleanRuleBody(CB), "[", "{", , , vbTextCompare)
                                                                    tempCleanRuleBody(CB) = Replace(tempCleanRuleBody(CB), "]", "}", , , vbTextCompare)
                                                          
                                                                End If
                                                                tempBody = tempBody & tempCleanRuleBody(CB)
                                                            Next CB
                                                            
                                                            
                                                            tempRuleBody = Split(tempBody, "[", , vbTextCompare)
                                                            If UBound(tempRuleBody) > 0 Then
                                                                tempFeature = Left(tempRuleBody(1), InStr(1, tempRuleBody(1), "]", vbTextCompare) - 1)
                                                            End If
                                                            add_RuleBody Body.Text, tempFeature, myLabel
                                                        
                                                            
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                ElseIf Rules.BaseName = "TableRules" Then
                                    For Each TableRuleDetails In Rules.ChildNodes
                                        spanstart = TableRuleDetails.Attributes.getnameditem("StartEvent").Text
                                        spanend = TableRuleDetails.Attributes.getnameditem("EndEvent").Text
                                        spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                        spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                        If requiredEffectivityDate = "" Then
                                            add_RuleDetails CoreXMLID, TableRuleDetails.Attributes.getnameditem("Code").Text, TableRuleDetails.Attributes.getnameditem("Description").Text, TableRuleDetails.Attributes.getnameditem("Intent").Text, TableRuleDetails.Attributes.getnameditem("StartEvent").Text, TableRuleDetails.Attributes.getnameditem("EndEvent").Text, TableRuleDetails.Attributes.getnameditem("IsEnabled").Text, TableRuleDetails.Attributes.getnameditem("IsLocked").Text, "Table"
                                            Debug.Print TableRuleDetails.Attributes.getnameditem("Code").Text
                                            For Each TableRule In TableRuleDetails.ChildNodes
                                                If TableRule.BaseName = "Labels" Then
                                                    For Each Label In TableRule.ChildNodes
                                                        add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                    Next
                                                End If
                                                If TableRule.BaseName = "Columns" Then
                                                    colFam = ""
                                                    For Each Family In TableRule.ChildNodes
                                                        add_RuleFamiles Family.Attributes.getnameditem("Code").Text
                                                        colFam = colFam & Family.Attributes.getnameditem("Code").Text & "|"
                                            
                                                    Next
                                                    colFam = Left(colFam, Len(colFam) - 1)
                                                End If
                                                If TableRule.BaseName = "Rows" Then
                                                    rowFam = ""
                                                    For Each Family In TableRule.ChildNodes
                                                        add_RuleFamiles Family.Attributes.getnameditem("Code").Text
                                                        rowFam = rowFam & Family.Attributes.getnameditem("Code").Text & "|"
                                            
                                                    Next
                                                    rowFam = Left(rowFam, Len(rowFam) - 1)
                                                End If
                                                If TableRule.BaseName = "Body" And impTableRules = True Then
                                                    
                                                    For Each tblheader In TableRule.ChildNodes
                                                        If tblheader.BaseName = "Row" Then
                                                            RowFeature = tblheader.Attributes.getnameditem("RowFeatureFamily").Text & "." & tblheader.Attributes.getnameditem("RowFeature").Text
                                                        End If
                                                        
                                                        For Each tCells In tblheader.ChildNodes
                                                            If tCells.BaseName = "Cells" Then
                                                                For Each tCell In tCells.ChildNodes
                                                                   
                                                                    
                                                                    For Each tcellcol In tCell.ChildNodes
                                                                        columnFeatures = ""
                                                                        For Each tCellFeat In tcellcol.ChildNodes
                                                                            columnFeatures = columnFeatures & tCellFeat.Attributes.getnameditem("FamilyCode").Text & "." & tCellFeat.Attributes.getnameditem("Code").Text & "|"
                                                                        Next
                                                                    
                                                                        columnFeatures = Left(columnFeatures, Len(columnFeatures) - 1)
                                                                        
                                                                        print_table_offerings RowFeature, columnFeatures, rowFam, colFam
                                                                    Next
                                                                    
                                                                    
                                                                Next
                                                            
                                                            End If
                                                        Next tCells
                                                        
                                                    Next
                                                
                                                End If
                                            Next
                                        Else
                                    
                                            If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                add_RuleDetails CoreXMLID, TableRuleDetails.Attributes.getnameditem("Code").Text, TableRuleDetails.Attributes.getnameditem("Description").Text, TableRuleDetails.Attributes.getnameditem("Intent").Text, TableRuleDetails.Attributes.getnameditem("StartEvent").Text, TableRuleDetails.Attributes.getnameditem("EndEvent").Text, TableRuleDetails.Attributes.getnameditem("IsEnabled").Text, TableRuleDetails.Attributes.getnameditem("IsLocked").Text, "Table"
                                                For Each TableRule In TableRuleDetails.ChildNodes
                                                    If TableRule.BaseName = "Labels" Then
                                                        For Each Label In TableRule.ChildNodes
                                                            add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                        Next
                                                    End If
                                                    If TableRule.BaseName = "Columns" Then
                                                        colFam = ""
                                                        For Each Family In TableRule.ChildNodes
                                                            add_RuleFamiles Family.Attributes.getnameditem("Code").Text
                                                            colFam = colFam & Family.Attributes.getnameditem("Code").Text & "|"
                                                
                                                        Next
                                                        colFam = Left(colFam, Len(colFam) - 1)
                                                    End If
                                                    If TableRule.BaseName = "Rows" Then
                                                        rowFam = ""
                                                        For Each Family In TableRule.ChildNodes
                                                            add_RuleFamiles Family.Attributes.getnameditem("Code").Text
                                                            rowFam = rowFam & Family.Attributes.getnameditem("Code").Text & "|"
                                                
                                                        Next
                                                        rowFam = Left(rowFam, Len(rowFam) - 1)
                                                    End If
                                                    If TableRule.BaseName = "Body" And impTableRules = True Then
                                                    
                                                        For Each tblheader In TableRule.ChildNodes
                                                            If tblheader.BaseName = "Row" Then
                                                                RowFeature = tblheader.Attributes.getnameditem("RowFeatureFamily").Text & "." & tblheader.Attributes.getnameditem("RowFeature").Text
                                                            End If
                                                            
                                                            For Each tCells In tblheader.ChildNodes
                                                                If tCells.BaseName = "Cells" Then
                                                                    For Each tCell In tCells.ChildNodes
                                                                       
                                                                        
                                                                        For Each tcellcol In tCell.ChildNodes
                                                                            columnFeatures = ""
                                                                            For Each tCellFeat In tcellcol.ChildNodes
                                                                                columnFeatures = columnFeatures & tCellFeat.Attributes.getnameditem("FamilyCode").Text & "." & tCellFeat.Attributes.getnameditem("Code").Text & "|"
                                                                            Next
                                                                        
                                                                            columnFeatures = Left(columnFeatures, Len(columnFeatures) - 1)
                                                                            
                                                                            print_table_offerings RowFeature, columnFeatures, rowFam, colFam
                                                                        Next
                                                                        
                                                                        
                                                                    Next
                                                                
                                                                End If
                                                            Next tCells
                                                            
                                                        Next
                                                    
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                ElseIf Rules.BaseName = "PackRules" Then
                                    For Each PackRuleDetails In Rules.ChildNodes
                                        spanstart = PackRuleDetails.Attributes.getnameditem("StartEvent").Text
                                        spanend = PackRuleDetails.Attributes.getnameditem("EndEvent").Text
                                        spanstartdate = Format(xmlEvents(spanstart), "YYYY/MM/DD HH:MM:SS")
                                        spansenddate = Format(xmlEvents(spanend), "YYYY/MM/DD HH:MM:SS")
                                        'If requiredEffectivityDate = "" Then
                                        If 1 = 1 Then 'Import all pack rules
                                            add_RuleDetails CoreXMLID, PackRuleDetails.Attributes.getnameditem("Code").Text, PackRuleDetails.Attributes.getnameditem("Description").Text, PackRuleDetails.Attributes.getnameditem("Intent").Text, PackRuleDetails.Attributes.getnameditem("StartEvent").Text, PackRuleDetails.Attributes.getnameditem("EndEvent").Text, PackRuleDetails.Attributes.getnameditem("IsEnabled").Text, PackRuleDetails.Attributes.getnameditem("IsLocked").Text, "Pack", PackRuleDetails.Attributes.getnameditem("PackFeature").Text, PackRuleDetails.Attributes.getnameditem("PackFeatureFamily").Text
                                            For Each packrule In PackRuleDetails.ChildNodes
                                                If packrule.BaseName = "Labels" Then
                                                    For Each Label In packrule.ChildNodes
                                                        add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                    Next
                                                End If
                                                If packrule.BaseName = "Body" Then
                                                    For Each packrulefeature In packrule.ChildNodes
                                                        If packrulefeature.ChildNodes.Length > 0 Then
                                                            For Each Condition In packrulefeature.ChildNodes
                                                                add_PackRuleFeatures packrulefeature.Attributes.getnameditem("Code").Text, packrulefeature.Attributes.getnameditem("FamilyCode").Text, Condition.Text
                                                            Next
                                                        Else
                                                            add_PackRuleFeatures packrulefeature.Attributes.getnameditem("Code").Text, packrulefeature.Attributes.getnameditem("FamilyCode").Text, ""
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        Else
                                            If spanstartdate <= requiredEffectivityDate And requiredEffectivityDate < spansenddate Then
                                                add_RuleDetails CoreXMLID, PackRuleDetails.Attributes.getnameditem("Code").Text, PackRuleDetails.Attributes.getnameditem("Description").Text, PackRuleDetails.Attributes.getnameditem("Intent").Text, PackRuleDetails.Attributes.getnameditem("StartEvent").Text, PackRuleDetails.Attributes.getnameditem("EndEvent").Text, PackRuleDetails.Attributes.getnameditem("IsEnabled").Text, PackRuleDetails.Attributes.getnameditem("IsLocked").Text, "Pack", PackRuleDetails.Attributes.getnameditem("PackFeature").Text, PackRuleDetails.Attributes.getnameditem("PackFeatureFamily").Text
                                                For Each packrule In PackRuleDetails.ChildNodes
                                                    If packrule.BaseName = "Labels" Then
                                                        For Each Label In packrule.ChildNodes
                                                            add_RuleLabels Label.Attributes.getnameditem("Name").Text
                                                        Next
                                                    End If
                                                    If packrule.BaseName = "Body" Then
                                                        For Each packrulefeature In packrule.ChildNodes
                                                            If packrulefeature.ChildNodes.Length > 0 Then
                                                                For Each Condition In packrulefeature.ChildNodes
                                                                    add_PackRuleFeatures packrulefeature.Attributes.getnameditem("Code").Text, packrulefeature.Attributes.getnameditem("FamilyCode").Text, Condition.Text
                                                                Next
                                                            Else
                                                                add_PackRuleFeatures packrulefeature.Attributes.getnameditem("Code").Text, packrulefeature.Attributes.getnameditem("FamilyCode").Text, ""
                                                            End If
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    End If
                    If Branch.BaseName = "SellableUnits" And impSU = True Then
                        For Each SellableUnits In Branch.ChildNodes
                            If SellableUnits.BaseName = "SellableUnit" Then

                                add_SellableUnit CoreXMLID, SellableUnits.Attributes.getnameditem("Code").Text, SellableUnits.Attributes.getnameditem("Description").Text, SellableUnits.Attributes.getnameditem("Active").Text
                                
                                For Each SellableUnit In SellableUnits.ChildNodes
                                    If SellableUnit.BaseName = "SellableUnitDefiningFeatures" Then
                                        For Each SellableUnitDefiningFeatures In SellableUnit.ChildNodes
                                            add_SellableUnitFeatures SUID, SellableUnitDefiningFeatures.Attributes.getnameditem("Code").Text, SellableUnitDefiningFeatures.Attributes.getnameditem("FamilyCode").Text
                                        Next
                                    End If
                                Next
                                
                            End If
                        Next
                    End If
                    If impSUOnly = False Then
                        If Branch.BaseName = "FeaturePriorities" Then
                            For Each FamilyDetails In Branch.ChildNodes
                                FamilyCode = FamilyDetails.Attributes.getnameditem("Code").Text
                                FOrder = 1
                                If FamilyDetails.BaseName = "Family" Then
    
            
                                    
                                    For Each FeatureList In FamilyDetails.ChildNodes
                                        If FeatureList.BaseName = "Features" Then
                                            For Each fPriority In FeatureList.ChildNodes
                                                FeatureCode = fPriority.Attributes.getnameditem("Code").Text
                                                'add_SellableUnitFeatures SUID, SellableUnitDefiningFeatures.Attributes.getnameditem("Code").Text, SellableUnitDefiningFeatures.Attributes.getnameditem("FamilyCode").Text
                                                
                                                updateSQL = ""
                                                updateSQL = "UPDATE coreFeatureAssociations "
                                                updateSQL = updateSQL & "set FeaturePriority =" & FOrder & " "
                                                updateSQL = updateSQL & "Where CoreXMLID=" & CoreXMLID & " AND FeatureCode='" & FeatureCode & "'"
                                                
                                                
                                                DoCmd.SetWarnings False
                                                DoCmd.RunSQL updateSQL
                                                DoCmd.SetWarnings True
                                                FOrder = FOrder + 1
                                            Next
                                        End If
                                    Next
                                    
                                End If
                            Next
                        End If
                    End If
                Next
            Next
        End If
    Next
    update_CoreXML_Details
    Set XDoc = Nothing
    Set dbs = Nothing
    
    Debug.Print "Ended: " & Now
    
End Sub

Function update_CoreXML_Details()
    Dim rstCoreXML As Recordset
    strsql = "select * from coreCoreXMLDetails where [Revision] = " & CoreXMLID
    Set rstCoreXML = dbs.OpenRecordset(strsql, dbOpenDynaset)

    
    rstCoreXML.MoveFirst
    With rstCoreXML
        .Edit
        .Fields("ImportEndDateTime") = Format(Now, "YYYY-MM-DD HH:MM:SS")
        .Fields("XML_Desc") = XMLDesc
        .Update
    End With
    
    Set rstcodexml = Nothing
End Function

Function add_CoreXML_Details(CoreXMLFile As String, PR As String, ImportEvent As String)

    Dim rstCoreXML As Recordset
check_XML_id:
    Set rstCoreXML = dbs.OpenRecordset("select * from coreCoreXMLDetails where [Revision] = " & CoreXMLID & " and PR = '" & PR & "'", dbOpenDynaset)

    If rstCoreXML.RecordCount > 0 Then
        If impSU = True Then
            impExists = True
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * from coreSellableUnits where [CoreXMLID] = " & CoreXMLID & ";"
            DoCmd.SetWarnings True
            Exit Function
        End If
        
        temp = MsgBox("The Revision of data in the Core XML file selected appears to already have ran.  Do you want to refresh the data?", vbQuestion + vbYesNo, "Refresh Data?")
        
        If temp = vbNo Then
            abort_import = True
            Exit Function
        Else
            abort_import = False
            DoCmd.SetWarnings False
            DoCmd.SetWarnings False
            
            Dim rstRules As Recordset
            Dim rstRuleLabels As Recordset
            Set rstRules = dbs.OpenRecordset("select * from coreRules where corexmlid=" & CoreXMLID)
            If rstRules.RecordCount = 0 Then GoTo skip_labels
            rstRules.MoveFirst
            
            While Not rstRules.EOF
                RuleID = rstRules.Fields("RuleLink")
                Set rstRuleLabels = rstRules.Fields("Labels").value
                
                If rstRuleLabels.RecordCount > 0 Then
                    While Not rstRuleLabels.EOF
                        rstRuleLabels.Delete
                        rstRuleLabels.MoveNext
                    Wend
                    
                End If
                
                
                rstRules.MoveNext
            Wend
skip_labels:
            'DoCmd.RunSQL "DELETE coreRules.Labels.Value FROM coreRules WHERE coreRules.CoreXMLID= " & XMLID & ";"
            'DoCmd.RunSQL "DELETE * FROM coreFeatureAssociations where [CoreXMLID] = " & XMLID & ";"
            'DoCmd.RunSQL "DELETE * FROM coreFormatedFA where [CoreXMLID] = " & XMLID & ";"
            'DoCmd.RunSQL "DELETE * FROM coreMarketAvailability where [CoreXMLID] = " & XMLID & ";"
            DoCmd.RunSQL "DELETE * from coreRules where [CoreXMLID] = " & CoreXMLID & ";"
            DoCmd.RunSQL "DELETE * from coreSellableUnits where [CoreXMLID] = " & CoreXMLID & ";"
            DoCmd.RunSQL "DELETE * FROM coreCoreXMLDetails where [Revision] = " & CoreXMLID & ";"
            'DoCmd.RunSQL "DELETE * FROM coreBrochureModels where [Revision] = " & XMLID & ";"
            'DoCmd.RunSQL "DELETE * FROM coreProgrammes where [Revision] = " & XMLID & ";"
            DoCmd.SetWarnings True
            
            rstRules.Close
            
            Set rstRuleLabels = Nothing
            Set rstRules = Nothing
        End If
    End If

    If DCount("*", "coreCoreXMLDetails", "Revision=" & CoreXMLID) > 0 Then
        CoreXMLID = CoreXMLID + 1
        GoTo check_XML_id
    End If
    
    A = FileDateTime(CoreXMLFile)
    
    With rstCoreXML
        .AddNew
        .Fields("Revision") = CoreXMLID
        .Fields("Branch") = CoreXMLBranch
        .Fields("PR") = PR
        .Fields("Event") = ImportEvent
        .Fields("CoreXMLDetails") = CoreXMLFile
        .Fields("CoreXMLFileDateStamp") = A
        .Fields("ImportStartDateTime") = Format(Now, "YYYY-MM-DD HH:MM:SS")
        .Fields("XML_Desc") = XMLDesc
        .Update
    End With
    
    Set rstcodexml = Nothing

End Function

Function add_programme(CoreXMLID As Long, pCode As String, pDesc As String)

    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreProgrammes (CoreXMLID, ProgrammeCode, ProgrammeDesc) " & _
                 "VALUES (" & CoreXMLID & ", '" & pCode & "','" & pDesc & "')"
        If Debug_Mode = True Then Debug.Print strsql
        DoCmd.RunSQL strsql

    Else
    
        Dim rstProgramme As Recordset
        Set rstProgramme = dbs.OpenRecordset("coreProgrammes", dbOpenDynaset)
        
        With rstProgramme
            .AddNew
            .Fields("ProgrammeLink") = CoreXMLID & "-" & pCode
            .Fields("CoreXMLID") = CoreXMLID
            .Fields("ProgrammeCode") = pCode
            .Fields("ProgrammeDesc") = pDesc
            .Update
        End With
        
        Set rstProgramme = Nothing
    End If
End Function

Function create_xmlEventsevents_list()
    Set xmlEvents = New Scripting.Dictionary
    
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] Like '" & CoreXMLID & "*'", dbOpenDynaset)
        
    rstEvents.MoveFirst
    
    While Not rstEvents.EOF
        EventCode = rstEvents.Fields("EventCode")
        SequenceDate = rstEvents.Fields("SequenceDate")
        xmlEvents.Add EventCode, SequenceDate
        rstEvents.MoveNext
    Wend
    Set rstEvents = Nothing
End Function
Function update_event_sequence(CoreXMLID As Long, eCode As String, newDate)
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] Like '" & CoreXMLID & "*'", dbOpenDynaset)
        
    
    rstEvents.MoveFirst
    rstEvents.FindFirst "[EventCode] = '" & eCode & "'"
    With rstEvents
        .Edit
        .Fields("SequenceDate") = newDate
        .Update
    End With
    
End Function


Function add_events(CoreXMLID As Long, eCode As String, eDesc As String, eDate As String, eType As String, pCode As String, Optional eSequenceDate As String)

    Dim rstProgramme As Recordset
    
    
    If isSubModel = False Then
        Set rstProgramme = dbs.OpenRecordset("coreProgrammes", dbOpenDynaset)
        With rstProgramme
          
            SearchC = "[ProgrammeCode] = '" & pCode & "' and[CoreXMLID] = " & CoreXMLID
            .FindFirst SearchC
            myProgrammeID = .Fields("ProgrammeLink").value
            myProgrammeCode = .Fields("ProgrammeCode").value
        End With
    
        Set rstProgramme = Nothing
    
    Else
        myProgrammeID = CoreXMLID & "-SUB"
        myProgrammeCode = "SUB"
        pCode = "SUB"
    End If

    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("coreEvents", dbOpenDynaset)
    
    With rstEvents
        .AddNew
        .Fields("EventLink") = CoreXMLID & "-" & pCode & "-" & eCode
        .Fields("ProgrammeID") = CoreXMLID & "-" & pCode
        .Fields("EventCode") = eCode
        .Fields("EventDesc") = eDesc
        eDate = Replace(eDate, "Z", "")
        eDate = Replace(eDate, "T", " ")
        .Fields("EventDate") = Format(eDate, "YYYY-MM-DD HH:MM:SS")
        .Fields("EventType") = eType
        If eSequenceDate <> "" Then
            eSequenceDate = Replace(eSequenceDate, "Z", "")
            eSequenceDate = Replace(eSequenceDate, "T", " ")
            .Fields("SequenceDate") = Format(eSequenceDate, "YYYY-MM-DD HH:MM:SS")
        End If
        .Update
        .MoveLast
    End With
    
    If eType = "VolumeIn" Then
        With rstEvents
            .AddNew
            .Fields("EventLink") = CoreXMLID & "-" & pCode & "-" & eCode & " OUT"
            .Fields("ProgrammeID") = CoreXMLID & "-" & pCode
            .Fields("EventCode") = eCode & " OUT"
            .Fields("EventDesc") = "Start of Programme"
            eDate = Replace(eDate, "Z", "")
            eDate = Replace(eDate, "T", " ")
            
            .Fields("EventDate") = Format(eDate, "YYYY-MM-DD HH:MM:SS")
            .Fields("EventType") = "VolumeOut"
            .Fields("SequenceDate") = Format(eDate, "YYYY-MM-DD HH:MM:SS")
            .Update
            .MoveLast
        End With
    End If
    
    Set rstEvents = Nothing

    
    
    eventlists(e, 1) = eCode
    eventlists(e, 2) = eType
    eventlists(e, 3) = CoreXMLID & "-" & pCode
    eventlists(e, 4) = Format(eDate, "YYYY-MM-DD HH:MM:SS")
    e = e + 1
    
    
End Function


Function add_FeatureAssociations(CoreXMLID As Long, fCode As String, fFamily As String, fPriority As String, eStart As String, eEnd As String)

    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & eStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & eEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing
    
    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreFeatureAssociations (CoreXMLID, FeatureCode, FeatureFamily, FeaturePriority, StartEventID, EndEventID) " & _
                 "VALUES (" & CoreXMLID & ", '" & fCode & "','" & fFamily & "'," & fPriority & "," & StartEventID & "," & EndEventID & ")"
        If Debug_Mode = True Then Debug.Print strsql
        DoCmd.RunSQL strsql
    Else
        Dim rstFeatures As Recordset
        Set rstFeatures = dbs.OpenRecordset("coreFeatureAssociations", dbOpenDynaset)
        
        With rstFeatures
            .AddNew
            .Fields("CoreXMLID") = CoreXMLID
            .Fields("FeatureCode") = fCode
            .Fields("FeatureFamily") = fFamily
            .Fields("FeaturePriority") = fPriority
            .Fields("StartEventID") = StartEventID
            .Fields("EndEventID") = EndEventID
            .Update
        End With
        
        Set rstFeatures = Nothing
    End If
End Function

Function add_ModelYearMappings(CoreXMLID As Long, fCode As String, fFamily As String, eStart As String, eEnd As String)
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & eStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & eEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing
    
    
    Dim rstModelYears As Recordset
    Set rstModelYears = dbs.OpenRecordset("coreModelYearMappings", dbOpenDynaset)
    
    With rstModelYears
        .AddNew
        .Fields("MYLink") = CoreXMLID & "-" & fCode
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("MYFeatureCode") = fCode
        .Fields("MYFeatureFamily") = fFamily
        .Fields("MYStartEventID") = StartEventID
        .Fields("MYEndEventID") = EndEventID
        .Update
        .MoveLast
        MYID = .Fields("MYLink")
    End With
    
    Set rstModelYears = Nothing
End Function

Function add_BuildPhases(CoreXMLID As Long, fCode As String, fFamily As String, fType As String, eStart As String, eEnd As String)
    
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & eStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & eEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing
    
    Dim rstBuildPhases As Recordset
    Set rstBuildPhases = dbs.OpenRecordset("coreBuildPhaseMappings", dbOpenDynaset)
    
    With rstBuildPhases
        .AddNew
        .Fields("MYID") = MYID
        .Fields("BPFeatureCode") = fCode
        .Fields("BPFeatureFamily") = fFamily
        .Fields("BPType") = fType
        .Fields("BPStartEventID") = StartEventID
        .Fields("BPEndEventID") = EndEventID
        .Update
    End With
    
    Set rstBuildPhases = Nothing
    
End Function

Function add_BrochureModels(CoreXMLID As Long, fCode As String, fDesc As String, fFeature As String, fFamily As String, AADA, AAGA, PBHA, ZZUW)

    Dim rstBrochureModels As Recordset
    Set rstBrochureModels = dbs.OpenRecordset("coreBrochureModels", dbOpenDynaset)
    
    With rstBrochureModels
        .AddNew
        .Fields("BMCLink") = CoreXMLID & "-" & fCode
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("BMCCode") = fCode
        .Fields("BMCDesc") = fDesc
        .Fields("BMCFeatureCode") = fFeature
        .Fields("BMCFeatureFamily") = fFamily
        .Fields("Engine") = AAGA
        .Fields("Wheelbase") = ZZUW
        .Fields("WheelsDriven") = PBHA
        .Fields("BodyStyle") = AADA
        .Update
        .MoveLast
        BMCID = .Fields("BMCLink")
    End With
    
    Set rstBrochureModels = Nothing

End Function

Function add_Derivatives(CoreXMLID As Long, fCode As String, fFamily As String)

    Dim Derivatives As Recordset
    Set Derivatives = dbs.OpenRecordset("coreDerivatives", dbOpenDynaset)
    
    With Derivatives
        .AddNew
        .Fields("BMCID") = BMCID
        .Fields("DerFeatureCode") = fCode
        .Fields("DerFeatureFamily") = fFamily
        .Update
    End With
    
    Set Derivatives = Nothing

End Function

Function add_marketAvailability(CoreXMLID As Long, mCode As String, mAvailable As String, mStart As String, mEnd As String, BMCCode As String, DPACKCode As String)

    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & mStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & mEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing
    
    If Import_Mode = "DOCMD" Then
    
        strsql = "INSERT INTO coreMarketAvailability (CoreXMLID,Market,IsAvailable, MAStartEventID, MAEndEventID, MABrochureModel, MADerivative) " & _
                 "VALUES (" & CoreXMLID & ", '" & mCode & "', " & mAvailable & ", " & StartEventID & "," & EndEventID & ", '" & BMCCode & "', '" & DPACKCode & "')"
        If Debug_Mode = True Then Debug.Print strsql
        
        DoCmd.RunSQL strsql
        
    Else
    
        Dim rstMarketAvailability As Recordset
        Set rstMarketAvailability = dbs.OpenRecordset("coreMarketAvailability", dbOpenDynaset)
        
        With rstMarketAvailability
            .AddNew
            .Fields("CoreXMLID") = CoreXMLID
            .Fields("Market") = mCode
            .Fields("IsAvailable") = mAvailable
            .Fields("MAStartEventID") = StartEventID
            .Fields("MAEndEventID") = EndEventID
            .Fields("MABrochureModel") = BMCCode
            .Fields("MADerivative") = DPACKCode
            .Update
        End With
        
        Set rstMarketAvailability = Nothing
        
    End If
End Function


Function add_IncludedFamilies(CoreXMLID As Long, fCode As String, fStart As String, fEnd As String)
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & fStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value

    End With
        
    With rstEvents
        
        SearchC = "[EventCode] = '" & fEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value

    End With
        
    Dim IncludedFamilies As Recordset
    Set IncludedFamilies = dbs.OpenRecordset("coreFAFamilies", dbOpenDynaset)
    
    With IncludedFamilies
        .AddNew
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("FAFamilyCode") = fCode
        .Fields("FAIncStartEvent") = StartEventID
        .Fields("FAIncEndEvent") = EndEventID
        .Update
    End With
    
    Set IncludedFamilies = Nothing
End Function


Function add_FASpan(CoreXMLID As Long, fFamily As String, faStart As String, faEnd As String)
    Dim rstFASpan As Recordset
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & faStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & faEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing
    
    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreFeatureApplicability (CoreXMLID, FAFamily, FAStartEventID, FAEndEventID) " & _
                 "VALUES (" & CoreXMLID & ", '" & fFamily & "'," & StartEventID & "," & EndEventID & ")"
        If Debug_Mode = True Then Debug.Print strsql
       
        DoCmd.RunSQL strsql
        
        
        Set rstFASpan = dbs.OpenRecordset("SELECT ID from coreFeatureApplicability where [CoreXMLID] = " & CoreXMLID, dbOpenDynaset)
        rstFASpan.MoveLast
        FASpanID = rstFASpan.Fields("ID")
    Else
    
        
        Set rstFASpan = dbs.OpenRecordset("coreFeatureApplicability", dbOpenDynaset)
        
        With rstFASpan
            .AddNew
            .Fields("CoreXMLID") = CoreXMLID
            .Fields("FAFamily") = fFamily
            .Fields("FAStartEventID") = StartEventID
            .Fields("FAEndEventID") = EndEventID
            .Update
            .MoveLast
            FASpanID = .Fields("ID")
        End With
        
        
        
        Set rstFASpan = Nothing
    End If
End Function

Function add_FASpanMarket(Market As String)
    
    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreFASpanMarkets (FASpanID, MarketCode) " & _
                 "VALUES (" & FASpanID & ", '" & Market & "')"
        If Debug_Mode = True Then Debug.Print strsql
        DoCmd.RunSQL strsql
        
        
    Else
        Dim rstFASpanMarket As Recordset
        Set rstFASpanMarket = dbs.OpenRecordset("coreFASpanMarkets", dbOpenDynaset)
        
        With rstFASpanMarket
            .AddNew
            .Fields("FASpanID") = FASpanID
            .Fields("MarketCode") = Market
            .Update
        End With
        
        Set rstFASpanMarket = Nothing
    End If
End Function

Function add_FASpanBMD(BMC As String, DPACK As String)
    
    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreFASpanBrochureModelDerivatives (FASpanID, BrochureModel, Derivative) " & _
                 "VALUES (" & FASpanID & ", '" & BMC & "','" & DPACK & "')"
        If Debug_Mode = True Then Debug.Print strsql
        DoCmd.RunSQL strsql
    
    Else

        Dim rstFASpanBMD As Recordset
        Set rstFASpanBMD = dbs.OpenRecordset("coreFASpanBrochureModelDerivatives", dbOpenDynaset)
        
        With rstFASpanBMD
            .AddNew
            .Fields("FASpanID") = FASpanID
            .Fields("BrochureModel") = BMC
            .Fields("Derivative") = DPACK
            .Update
        End With
        
        Set rstFASpanBMD = Nothing
    End If
End Function

Function add_FASpanFeature(faFeature As String, faFamily As String, faAvailability As String, faMarketingValue As String)
    
    If Import_Mode = "DOCMD" Then
        strsql = "INSERT INTO coreFASpanFeature (FASpanID, FeatureCode, FamilyCode, Availability, MarketingValue) " & _
                 "VALUES (" & FASpanID & ", '" & faFeature & "','" & faFamily & "','" & faAvailability & "','" & faMarketingValue & "')"
        If Debug_Mode = True Then Debug.Print strsql
        DoCmd.RunSQL strsql
        
    Else
    
        Dim rstFASpanFeature As Recordset
        Set rstFASpanFeature = dbs.OpenRecordset("coreFASpanFeature", dbOpenDynaset)
        
        With rstFASpanFeature
            .AddNew
            .Fields("FASpanID") = FASpanID
            .Fields("FeatureCode") = faFeature
            .Fields("FamilyCode") = faFamily
            .Fields("Availability") = faAvailability
            .Fields("MarketingValue") = faMarketingValue
            .Update
        End With
        'Debug.Print rstFASpanFeature.RecordCount
        Set rstFASpanFeature = Nothing
    End If
End Function

Function combined_FA(sDate As String, eDate As String)
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & sDate & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & eDate & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Set rstEvents = Nothing

    Dim rstFACombi As Recordset
    Set rstFACombi = dbs.OpenRecordset("select * from coreFormatedFA where [CoreXMLID] = " & CoreXMLID, dbOpenDynaset)

    For Each Market In FA_market
        For Each BMC In FA_bmc
            tempBMC = Split(BMC, "|", , vbTextCompare)
            mybmc = tempBMC(0)
            mydpack = tempBMC(1)
            For Each feat In FA_feature
                tempFeat = Split(FA_feature(feat), "|", , vbTextCompare)
                myFam = tempFeat(0)
                myAvail = tempFeat(1)
                myValue = tempFeat(2)
                
                With rstFACombi
                    .AddNew
                    .Fields("CoreXMLID") = CoreXMLID
                    .Fields("MarketCode") = Market
                    .Fields("BrochureModel") = mybmc
                    .Fields("Derivative") = mydpack
                    .Fields("FAFamily") = myFam
                    .Fields("FeatureCode") = feat
                    .Fields("Availability") = myAvail
                    .Fields("MarketingValue") = myValue
                    .Fields("StartDate") = StartEventID
                    .Fields("EndDate") = EndEventID
                    .Update
                End With
            Next feat
        Next BMC
    Next Market
    
    
    Set rstFACombi = Nothing
End Function
Function add_RuleDetails(CoreXMLID As Long, rCode As String, rDesc As String, rIntent As String, rStart As String, rEnd As String, rEnabled As String, rLocked As String, rType As String, Optional pFeature As String, Optional pFamily As String)
    Dim rstEvents As Recordset
    Set rstEvents = dbs.OpenRecordset("select * from coreEvents where [ProgrammeID] like '" & CoreXMLID & "*'", dbOpenDynaset)
    
    With rstEvents
        
        SearchC = "[EventCode] = '" & rStart & "'"
        .FindFirst SearchC
        StartEventID = .Fields("EventLink").value
        
        SearchC = "[EventCode] = '" & rEnd & "'"
        .FindFirst SearchC
        EndEventID = .Fields("EventLink").value
    End With
    
    Dim rstRuleDetails As Recordset
    Set rstRuleDetails = dbs.OpenRecordset("coreRules", dbOpenDynaset)
    
    With rstRuleDetails
        .AddNew
        .Fields("RuleLink") = CoreXMLID & "-" & rCode
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("RuleCode") = rCode
        .Fields("RuleDescription") = Left(rDesc, 255)
        .Fields("RuleIntent") = rIntent
        .Fields("RuleStartEventID") = StartEventID
        .Fields("RuleEndEventID") = EndEventID
        .Fields("IsEnabled") = rEnabled
        .Fields("IsLocked") = rLocked
        .Fields("RuleType") = rType
        
        If rType = "Table" Then
            ruleDesc = rDesc
        
            If InStr(1, ruleDesc, "[", vbTextCompare) > 1 Then
                rulesplit = Split(ruleDesc, "[", , vbTextCompare)

                For A = 0 To UBound(rulesplit)
                    If Left(rulesplit(A), 9) = "Version: " Then

                        temp = Mid(rulesplit(A), 10, Len(rulesplit(A)) - 10)
                        temp = Replace(temp, "]", "", , , vbTextCompare)
                        .Fields("Version") = temp
                    ElseIf Left(rulesplit(A), 3) = "FCT" Then
                        temp = Replace(rulesplit(A), "FCT", "", , , vbTextCompare)
                        temp = Replace(temp, " ", "", , , vbTextCompare)
                        temp = Replace(temp, ":", "", , , vbTextCompare)
                        temp = Replace(temp, "]", "", , , vbTextCompare)
                        
                        .Fields("FCT") = temp
                        
                    ElseIf Left(rulesplit(A), 6) = "FCIM: " Then
                        temp = Mid(rulesplit(A), 7, Len(rulesplit(A)) - 8)
                        .Fields("FCT") = Mid(rulesplit(A), 7, Len(rulesplit(A)) - 7)
                    End If
                
                Next A

            End If
        
 
        
        End If
        
        .Fields("PackFeature") = pFeature
        .Fields("PackFeatureFamily") = pFamily
        .Update
        RuleID = CoreXMLID & "-" & rCode
    End With
    
    Set rstRuleDetails = Nothing
     
End Function

Function add_RuleLabels(rLabel As String)
    Dim rstLabels As Recordset
    Set rstLabels = dbs.OpenRecordset("select * from mstLabels where LabelName = '" & rLabel & "'", dbOpenDynaset)
    
    If rstLabels.RecordCount = 0 Then
        With rstLabels
            .AddNew
            .Fields("LabelName") = rLabel
            .Update
        End With
    End If
    Set rstLabels = Nothing
    
    
    Dim rstRuleDetails As Recordset
    Set rstRuleDetails = dbs.OpenRecordset("select labels,MainLabel from coreRules where rulelink = '" & RuleID & "'", dbOpenDynaset)
    Dim rstRuleDetailsLabels As Recordset
    
    With rstRuleDetails
        .Edit
        If InStr(1, rLabel, "TRANSL", vbTextCompare) = 0 And InStr(1, rLabel, "OXO Rules Exclusion", vbTextCompare) = 0 Then
            .Fields("MainLabel") = rLabel
        End If
        Set rstRuleDetailsLabels = rstRuleDetails.Fields("Labels").value
 
        On Error Resume Next
        With rstRuleDetailsLabels
            .AddNew
            .Fields(0) = rLabel
            .Update
        End With

       
       .Update
    End With
    On Error GoTo 0
    Set rstRuleDetails = Nothing
    Set rstRuleDetailsLabels = Nothing
    
End Function

Function add_RuleBody(rBody As String, rFeature As String, rLabel As String)
    Dim rConsumedFeatures As String
    Dim restFeatures  As String
    
    Dim rstRuleBody As Recordset
    Set rstRuleBody = dbs.OpenRecordset("coreRuleBody", dbOpenDynaset)
    temppaint = ""
    
    
    tempBody = Split(rBody, vbLf, , vbTextCompare)
    
    For A = 0 To UBound(tempBody)
        If Left(tempBody(A), 2) = "//" Or tempBody(A) = "" Then
            
        Else
            rlogic = rlogic & vbCrLf & tempBody(A)
        End If
    Next A
    
    Dim ConsumedFeatures As Scripting.Dictionary
    Set ConsumedFeatures = New Scripting.Dictionary
    
    Dim restrictionFeatures As Scripting.Dictionary
    Set restrictionFeatures = New Scripting.Dictionary
    
    rBody = Replace(rBody, vbLf, vbCrLf, , , vbTextCompare)
    rlogic = Replace(rlogic, "AnyOf (", "AnyOf(", , , vbTextCompare)
    rlogic = Replace(rlogic, "NoneOf (", "NoneOf(", , , vbTextCompare)
    rlogic = Replace(rlogic, "AllOf (", "AllOf(", , , vbTextCompare)
    rlogic = Replace(rlogic, " )", ")", , , vbTextCompare)
    rlogic = Replace(rlogic, "( ", "(", , , vbTextCompare)
    rlogic = Replace(rlogic, " ,", ",", , , vbTextCompare)
    rlogic = Replace(rlogic, ", ", ",", , , vbTextCompare)
    rlogic = Replace(rlogic, ",", ", ", , , vbTextCompare)
    
    If temppaint <> "" And Not IsNull(temppaint) Then
        rConsumedFeatures = temppaint
    Else
        rConsumedFeatures = rFeature
    End If
    
    'If rFeature = "088TV" Then Stop
    findthen = InStr(1, rlogic, " then ", vbTextCompare)
    findAnd = InStr(1, rlogic, " and ", vbTextCompare)
    
    

    
    For Each feat In rMFDFeatureToFamily
        
        tempcheck = ""
        If InStr(1, rlogic, "[" & feat & "]", vbTextCompare) > 0 And InStr(1, rConsumedFeatures, feat, vbTextCompare) = 0 Then

            
            If rMFDDimension.Exists(feat) = True Then
                restrictionFeatures.Add rMFDFeatureToFamily(feat) & ".[" & feat & "]", 1

            ElseIf InStr(1, rlogic, feat, vbTextCompare) > 0 Then
                If InStr(1, rlogic, rFeature, vbTextCompare) < InStr(1, rlogic, feat, vbTextCompare) And InStr(1, rlogic, " then ", vbTextCompare) > InStr(1, rlogic, feat, vbTextCompare) Then
                    
                    
					restrictionFeatures.Add rMFDFeatureToFamily(feat) & ".[" & feat & "]", 1
					ConsumedFeatures.Add feat, 1
                    
                ElseIf InStr(1, rlogic, " then ", vbTextCompare) > InStr(1, rlogic, feat, vbTextCompare) = 0 Then
                    ConsumedFeatures.Add feat, 2
                Else
                     ConsumedFeatures.Add feat, 1
                End If
            Else

                
                    
                ConsumedFeatures.Add feat, 1
                
            End If
         
        End If
    
    Next feat

   

    Set ConsumedFeatures = modReasonsForRules.SortDictionaryByKey(ConsumedFeatures, xlAscending)
    Set restrictionFeatures = modReasonsForRules.SortDictionaryByKey(restrictionFeatures, xlAscending)
        
    
    
    For Each feat In ConsumedFeatures
        If InStr(1, rConsumedFeatures, feat, vbTextCompare) = 0 Then rConsumedFeatures = rConsumedFeatures & "," & feat
    Next feat
    
    restFeatures = ""
    For Each feat In restrictionFeatures
        restFeatures = restFeatures & "," & feat
    Next feat
    
    If Len(restFeatures) > 0 Then restFeatures = Right(restFeatures, Len(restFeatures) - 1)
    
    If rLabel <> "OXO Rules" Then GoTo got_FS_ID
    
    Dim fsid As Integer
    Dim rstFS As Recordset
     
     'If rFeature = "301MU" Then Stop
     
    fsid = get_FS_ID(rConsumedFeatures)
    If restFeatures = "" Or IsNull(restFeatures) Then
        RFID = 0
    Else
         RFID = get_FS_ID(restFeatures)
    End If
    
    Debug.Print rBody

got_FS_ID:
    

    With rstRuleBody
        .AddNew
        .Fields("RuleID") = RuleID
        .Fields("RuleDetails") = rBody
        If temppaint = "" Or IsNull(temppaint) Then
            .Fields("RuleFeature") = rFeature
        Else
            .Fields("RuleFeature") = temppaint
        End If
        .Fields("RuleLogic") = rlogic
        If rLabel = "OXO Rules" Then
            .Fields("ConsumedFeatures") = rConsumedFeatures
            .Fields("FS_ID") = fsid
            .Fields("RF_ID") = RFID
        End If
        .Update
    End With
    
    Set rstRuleBody = Nothing
End Function

Function add_PackRuleFeatures(PRFeature, prFamily, prCondition)
    Dim rstRuleFeatures As Recordset
    Set rstRuleFeatures = dbs.OpenRecordset("corePackRuleFeatures", dbOpenDynaset)
    
    With rstRuleFeatures
        .AddNew
        .Fields("RuleID") = RuleID
        .Fields("FeatureCode") = PRFeature
        .Fields("FamilyCode") = prFamily
        .Fields("Condition") = prCondition
        .Update
    End With
    
    Set rstRuleFeatures = Nothing
End Function

Function add_RuleFamiles(prFamily)
    Dim rstRuleFamiles As Recordset
    Set rstRuleFamiles = dbs.OpenRecordset("coreRuleFamilies", dbOpenDynaset)
    
    With rstRuleFamiles
        .AddNew
        .Fields("RuleID") = RuleID
        .Fields("FamilyCode") = prFamily
        .Update
    End With
    
    Set rstRuleFamiles = Nothing
End Function

Function add_SellableUnit(CoreXMLID As Long, SUCode As String, SUDesc As String, SUActive As String)
    Dim rstSellableUnits As Recordset
    Set rstSellableUnits = dbs.OpenRecordset("coreSellableUnits", dbOpenDynaset)
    
    With rstSellableUnits
        .AddNew
        .Fields("SULink") = CoreXMLID & "-" & SUCode
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("SUCode") = SUCode
        .Fields("SUDesc") = SUDesc
        .Fields("SUActive") = SUActive
        .Update
        SUID = CoreXMLID & "-" & SUCode
    End With
    
    Set rstSellableUnits = Nothing
End Function

Function add_SellableUnitFeatures(SUID As String, SUFeature As String, SUFamily As String)
    Dim rstSellableUnitFeatures As Recordset
    Set rstSellableUnitFeatures = dbs.OpenRecordset("coreSellableUnitFeatures", dbOpenDynaset)
    
    With rstSellableUnitFeatures
        .AddNew
        .Fields("SUID") = SUID
        .Fields("SUFeature") = SUFeature
        .Fields("SUFamily") = SUFamily
        .Update
    End With
    
    Set rstSellableUnitFeatures = Nothing
End Function

Function add_basics(PRCode, PRDesc, PRBrand, PRFeature)
    Dim rstBasics As Recordset
    Set rstBasics = dbs.OpenRecordset("coreBasics", dbOpenDynaset)
    
    With rstBasics
        .AddNew
        .Fields("CoreXMLID") = CoreXMLID
        .Fields("PRCode") = PRCode
        .Fields("PRDesc") = PRDesc
        .Fields("PRBrand") = PRBrand
        .Fields("PRFeature") = PRFeature
        .Update
    End With
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE coreCoreXMLDetails SET PR='" & PRCode & "' WHERE Revision=" & CoreXMLID
    DoCmd.SetWarnings True
    
    Set rstBasics = Nothing
End Function


Function refresh_library()
    
    Set XDocMani = CreateObject("Msxml2.DOMDocument.6.0")
    XDocMani.async = False: XDocMani.validateOnParse = False
    XDocMani.Load (ManifestXMLFile)
    
    Set manilists = XDocMani.DocumentElement
    
    For Each maniDetails In manilists.ChildNodes
        If maniDetails.BaseName = "Metadata" Then
            For Each manidetail In maniDetails.ChildNodes
                If manidetail.Attributes.getnameditem("Key").Text = "Branch" Then
                    CoreXMLBranch = manidetail.Attributes.getnameditem("Value").Text
                ElseIf manidetail.Attributes.getnameditem("Key").Text = "Revision" Then
                    CoreXMLID = manidetail.Attributes.getnameditem("Value").Text
                ElseIf manidetail.Attributes.getnameditem("Key").Text = "ExportDate" Then
                    coreXMLDate = manidetail.Attributes.getnameditem("Value").Text
                End If
            Next
        End If
    Next
    
    Set XDocLib = CreateObject("Msxml2.DOMDocument.6.0")
    XDocLib.async = False: XDocLib.validateOnParse = False
    XDocLib.Load (LibraryXML)
    On Error Resume Next
    DoCmd.SetWarnings False
    Set dbs = CurrentDb
    Set Librarylists = XDocLib.DocumentElement
    
    For Each librarydetails In Librarylists.ChildNodes
        If librarydetails.BaseName = "Library" Then
            For Each library In librarydetails.ChildNodes
                If library.BaseName = "Families" Then
                    DoCmd.RunSQL "DELETE * FROM mstFamilies;"
                    Dim rstFamilies As Recordset
                    Set rstFamilies = dbs.OpenRecordset("mstFamilies", dbOpenDynaset)
                    
                    For Each Families In library.ChildNodes
                        If Families.BaseName = "FeatureFamilies" Then
                            For Each Family In Families.ChildNodes
                                myPurpose = False
                                For Each FamilyDetails In Family.ChildNodes
                                    If FamilyDetails.BaseName = "PropertyValues" Then
                                        For Each PropertyValues In FamilyDetails.ChildNodes
                                            'Debug.Print PropertyValues.Attributes.getnameditem("Name").Text, PropertyValues.Attributes.getnameditem("Value").Text
                                            If PropertyValues.Attributes.getnameditem("Name").Text = "EnoviaPurpose" And PropertyValues.Attributes.getnameditem("Value").Text = "OXO" Then
                                                myPurpose = True
                                            End If
                                        Next PropertyValues
                                    End If
                                Next FamilyDetails

                                With rstFamilies
                                    .AddNew
                                    .Fields("FamilyCode") = Family.Attributes.getnameditem("Code").Text
                                    .Fields("FamilyDesc") = Family.Attributes.getnameditem("Description").Text
                                    .Fields("FeatureFamilyType") = Family.Attributes.getnameditem("FeatureFamilyType").Text
                                    .Fields("LessFeature") = Family.Attributes.getnameditem("LessFeature").Text
                                    .Fields("Lifecycle") = Family.Attributes.getnameditem("Lifecycle").Text
                                    .Fields("OXO_Purpose") = myPurpose
                                    .Update
                                End With
                                
                                
                            Next
                        End If
                    Next
                    Set rstFamilies = Nothing
                End If
                
                If library.BaseName = "Features" Then
                    
                    DoCmd.RunSQL "DELETE * FROM mstFeatures;"
                    Dim rstFeatures As Recordset
                    Set rstFeatures = dbs.OpenRecordset("mstFeatures", dbOpenDynaset)
                    
                    For Each Features In library.ChildNodes
                        If Features.BaseName = "Feature" Then
                            'For Each feature In Features.childnodes
                            myLRDesc = ""
                            myJagDesc = ""
                            For Each FeaturesDetails In Features.ChildNodes
                                If FeaturesDetails.BaseName = "Descriptions" Then
                                    For Each Descriptions In FeaturesDetails.ChildNodes

                                        If Descriptions.Attributes.getnameditem("Brand").Text = "andover" Then
                                            myLRDesc = Descriptions.Text
                                        ElseIf Descriptions.Attributes.getnameditem("Brand").Text = "jauar" Then
                                            myJagDesc = Descriptions.Text
                                        End If
                                    Next Descriptions
                                End If
                            Next FeaturesDetails
                            
                            With rstFeatures
                                .AddNew
                                .Fields("FeatureCode") = Features.Attributes.getnameditem("Code").Text
                                .Fields("FeatureDesc") = Features.Attributes.getnameditem("Description").Text
                                .Fields("Lifecycle") = Features.Attributes.getnameditem("Lifecycle").Text
                                .Fields("Family") = Features.Attributes.getnameditem("FeatureFamily").Text
                                .Fields("LRDesc") = myLRDesc
                                .Fields("JagDesc") = myJagDesc
                                .Update
                            End With
                                
                            
                        End If
                    Next
                    
                    Set rstFeatures = Nothing
                End If
            Next
        End If
    Next
    'Debug.Print coreXMLDate
    tempcoreXMLDate = Replace(coreXMLDate, "Z", "")
    tempcoreXMLDate = Replace(tempcoreXMLDate, "T", " ")
    tempcoreXMLDate = Format(Left(tempcoreXMLDate, 19), "DD/MM/YYYY HH:MM:SS")
    'If IsDate(tempcoreXMLDate) = True Then Stop
    'Debug.Print tempcoreXMLDate
    'Stop
    strsql = "UPDATE mstLastUpdated SET DateLastUpdated='" & tempcoreXMLDate & "' WHERE Type='Ace Library'"
    DoCmd.RunSQL strsql
    
    DoCmd.SetWarnings True
End Function

Sub test()
    GroupsXML = "C:\Users\sjones76\Documents\CMT\Test\Core XML to OXO\Library\msc.xml"
    refresh_Groups
End Sub
Function refresh_Groups()

Set XDocMani = CreateObject("Msxml2.DOMDocument.6.0")
    XDocMani.async = False: XDocMani.validateOnParse = False
    XDocMani.Load (ManifestXMLFile)
    
    Set manilists = XDocMani.DocumentElement
    
    For Each maniDetails In manilists.ChildNodes
        If maniDetails.BaseName = "Metadata" Then
            For Each manidetail In maniDetails.ChildNodes
                If manidetail.Attributes.getnameditem("Key").Text = "Branch" Then
                    CoreXMLBranch = manidetail.Attributes.getnameditem("Value").Text
                ElseIf manidetail.Attributes.getnameditem("Key").Text = "Revision" Then
                    CoreXMLID = manidetail.Attributes.getnameditem("Value").Text
                ElseIf manidetail.Attributes.getnameditem("Key").Text = "ExportDate" Then
                    coreXMLDate = manidetail.Attributes.getnameditem("Value").Text
                End If
            Next
        End If
    Next
    
    Set XDocLib = CreateObject("Msxml2.DOMDocument.6.0")
    XDocLib.async = False: XDocLib.validateOnParse = False
    XDocLib.Load (GroupsXML)
    On Error Resume Next
    DoCmd.SetWarnings False
    Set GroupsList = XDocLib.DocumentElement
    Set dbs = CurrentDb
    DoCmd.RunSQL "DELETE * FROM mstVistaGroups;"
    DoCmd.RunSQL "DELETE * FROM mstVistaGroupDetails;"
    Dim rstVistaGroups As Recordset
    Set rstVistaGroups = dbs.OpenRecordset("mstVistaGroups", dbOpenDynaset)
    Dim rstVistaGroupDetails As Recordset
    Set rstVistaGroupDetails = dbs.OpenRecordset("mstVistaGroupDetails", dbOpenDynaset)
    
    For Each aModules In GroupsList.ChildNodes
        For Each msc In aModules.ChildNodes
            For Each Groupsdetails In msc.ChildNodes
                If Groupsdetails.BaseName = "FeatureGroups" Then
                    For Each FeatureGroups In Groupsdetails.ChildNodes
                        Debug.Print FeatureGroups.Attributes.getnameditem("Code").Text, FeatureGroups.Attributes.getnameditem("Type").Text
                        
                        
                        If FeatureGroups.Attributes.getnameditem("Type").Text = "DisplayGroup" Then
                            With rstVistaGroups
                                .AddNew
                                .Fields("Code") = FeatureGroups.Attributes.getnameditem("Code").Text
                                .Fields("Type") = FeatureGroups.Attributes.getnameditem("Type").Text
                                .Fields("Description") = FeatureGroups.Attributes.getnameditem("Description").Text
                                .Update
                            End With
                            
                            For Each featureGroup In FeatureGroups.ChildNodes
                                
                                If featureGroup.BaseName = "Families" Then
                                    For Each Family In featureGroup.ChildNodes
                                        Debug.Print Families.Attributes.getnameditem("Code").Text
                                        With rstVistaGroupDetails
                                            .AddNew
                                            .Fields("Group") = FeatureGroups.Attributes.getnameditem("Code").Text
                                            .Fields("Family") = Family.Attributes.getnameditem("Code").Text
                                            .Update
                                        End With
                                    Next
                                
                                End If
                            Next
                            
                        End If
                    Next
                End If
            Next
        Next
    Next
    
    tempcoreXMLDate = Replace(coreXMLDate, "Z", "")
    tempcoreXMLDate = Replace(tempcoreXMLDate, "T", " ")
    tempcoreXMLDate = Format(Left(tempcoreXMLDate, 19), "DD/MM/YYYY HH:MM:SS")
    
    strsql = "UPDATE mstLastUpdated SET DateLastUpdated='" & tempcoreXMLDate & "' WHERE Type='Display Groups'"
    DoCmd.RunSQL strsql
    
    Set rstVistaGroups = Nothing
    Set rstVistaGroupDetails = Nothing
    DoCmd.SetWarnings True
End Function


Function get_feature_library()

    Dim rstRMFD As DAO.Recordset
    
    Set rstRMFD = dbs.OpenRecordset("SELECT * from tblMFD")
    Set rMFDFeatureDesc = New Scripting.Dictionary
    Set rMFDFeatureToFamily = New Scripting.Dictionary
    Set rMFDFeatureLRDesc = New Scripting.Dictionary
    Set rMFDFeatureJagDesc = New Scripting.Dictionary
    Set rMFDDimension = New Scripting.Dictionary
    
    rstRMFD.MoveFirst
    While Not rstRMFD.EOF
    
        If InStr(1, rstRMFD("System Description").value, chr(160), vbTextCompare) > 0 Then
            tempDesc = rstRMFD("System Description").value
            tempDesc = Replace(tempDesc, chr(160), chr(32), , , vbTextCompare)
            rMFDFeatureDesc.Add rstRMFD.Fields("Feature Code").value, tempDesc
        Else
            rMFDFeatureDesc.Add rstRMFD.Fields("Feature Code").value, rstRMFD("System Description").value
        End If

        rMFDFeatureToFamily.Add rstRMFD.Fields("Feature Code").value, rstRMFD("Family Code").value
        If InStr(1, rstRMFD("andoverDesc").value, chr(160), vbTextCompare) > 0 Then
            tempDesc = rstRMFD("andoverDesc").value
            tempDesc = Replace(tempDesc, chr(160), chr(32), , , vbTextCompare)
            rMFDFeatureLRDesc.Add rstRMFD.Fields("Feature Code").value, tempDesc
        Else
            rMFDFeatureLRDesc.Add rstRMFD.Fields("Feature Code").value, rstRMFD("andoverDesc").value
        End If
        
        If InStr(1, rstRMFD("jauarDesc").value, chr(160), vbTextCompare) > 0 Then
            tempDesc = rstRMFD("jauarDesc").value
            tempDesc = Replace(tempDesc, chr(160), chr(32), , , vbTextCompare)
            rMFDFeatureJagDesc.Add rstRMFD.Fields("Feature Code").value, tempDesc
        Else
            rMFDFeatureJagDesc.Add rstRMFD.Fields("Feature Code").value, rstRMFD("jauarDesc").value
        End If
        
        
        If rMFDDimension.Exists(rstRMFD.Fields("Feature Code").value) = False And (rstRMFD.Fields("Family Code").value = "MSCS" Or rstRMFD.Fields("Feature Family Type").value = "BMOD" Or rstRMFD.Fields("Feature Family Type").value = "DPCK" Or rstRMFD.Fields("Feature Family Type").value = "MRKT") Then
            rMFDDimension.Add rstRMFD.Fields("Feature Code").value, rstRMFD.Fields("Feature Family Type").value
        End If
        rstRMFD.MoveNext
    Wend
End Function


Function print_table_offerings(RowDetails, ColDetails, rowFams, colFams)

    Dim rstRuleFeatures As Recordset
    Set rstRuleFeatures = dbs.OpenRecordset("coreRuleTableCombinations", dbOpenDynaset)
    
    With rstRuleFeatures
        .AddNew
        .Fields("RuleID") = RuleID
        .Fields("RowFitment") = RowDetails
        .Fields("ColumnFitment") = ColDetails
        .Fields("rowFams") = rowFams
        .Fields("colFams") = colFams
        .Update
    End With
    
    Set rstRuleFeatures = Nothing


End Function


