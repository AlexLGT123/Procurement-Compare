/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleGenAI, Type } from "@google/genai";
import * as React from 'react';
// FIX: Corrected import syntax for react-dom/client.
import * as ReactDOM from 'react-dom/client';

const { jsPDF } = (window as any).jspdf;
const html2canvas = (window as any).html2canvas;
const XLSX = (window as any).XLSX;

// New interface for per-supplier summary points
interface SummaryPoint {
    fileName: string;
    points: string[];
}

// Interface for the structured summary
interface ExecutiveSummary {
    recommendation: string;
    strengths: SummaryPoint[];
    weaknesses: SummaryPoint[];
    negotiationPoints: SummaryPoint[];
}

// Nested interfaces for categorized scores
interface ComparisonQualityScores {
    CompletionScore: number;
    PrecisionScore: number;
}

interface QuotePerformanceScores {
    PriceCompetitivenessScore: number;
    DeliveryServiceScore: number;
}

// Interface for an individual quote's performance
interface QuotePerformance {
    fileName: string;
    scores: QuotePerformanceScores;
}

// New interfaces for data correction
type ItemId = number;
type GroupId = number;

interface ParsedItem {
    id: ItemId;
    groupId: GroupId;
    [key: string]: string | number;
}

interface ParsedData {
    [fileName: string]: ParsedItem[];
}

interface IssueDetail {
    issueType: keyof ComparisonQualityScores;
    fileName: string;
    itemId: ItemId;
    groupId: GroupId;
    explanation: string;
    problematicField?: string;
}

// Interface for type safety on the API response
interface ComparisonResult {
    comparisonQuality: ComparisonQualityScores;
    detailedIssues: IssueDetail[];
    parsedData: ParsedData;
    quotePerformance?: QuotePerformance[];
    executiveSummary?: ExecutiveSummary;
}

// Props interface for FileUpload component
interface FileUploadProps {
    files: File[];
    onAddFiles: (newFiles: File[]) => void;
    onRemoveFile: (fileName: string) => void;
    title: string;
    subtitle: string;
    placeholderText: string;
}

const FileUpload: React.FC<FileUploadProps> = React.memo(({ files, onAddFiles, onRemoveFile, title, subtitle, placeholderText }) => {
    const [isDragging, setIsDragging] = React.useState(false);
    const inputRef = React.useRef<HTMLInputElement | null>(null);

    const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(true);
    };

    const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(false);
    };

    const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            // FIX: Explicitly type `file` as File to resolve properties like 'type' and 'name'.
            const validFiles = Array.from(e.dataTransfer.files).filter((file: File) => 
                file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
                file.type === "application/vnd.ms-excel" ||
                file.type === "text/csv" ||
                file.name.endsWith('.xlsx') ||
                file.name.endsWith('.xls') ||
                file.name.endsWith('.csv')
            );
            if (validFiles.length > 0) {
              onAddFiles(validFiles);
            }
        }
    };

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files && e.target.files.length > 0) {
            onAddFiles(Array.from(e.target.files));
            if (inputRef.current) {
                inputRef.current.value = '';
            }
        }
    };

    return (
        <div className="upload-area">
            <h2 className="panel-title">{title}</h2>
             <p className="panel-subtitle">{subtitle}</p>
            <div
                className={`file-uploader ${isDragging ? "drag-over" : ""}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => inputRef.current?.click()}
            >
                <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
                <p>{placeholderText}</p>
                <input
                    type="file"
                    ref={inputRef}
                    onChange={handleChange}
                    style={{ display: "none" }}
                    accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv,.xlsx,.xls,.csv"
                    multiple
                />
            </div>
             <div className="file-list">
                {files.map(file => (
                    <div className="file-card" key={file.name}>
                        <svg className="file-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><line x1="10" y1="9" x2="8" y2="9"></line></svg>
                        <span className="file-name">{file.name}</span>
                        <button
                            className="remove-file-btn"
                            onClick={(e) => { e.stopPropagation(); onRemoveFile(file.name); }}
                            aria-label={`Remove ${file.name}`}
                        >
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                        </button>
                    </div>
                ))}
            </div>
        </div>
    );
});

interface ScoreCardProps {
    value: number;
    label: string;
    description: string;
    isActionable?: boolean;
    onClick?: () => void;
}

const ScoreCard: React.FC<ScoreCardProps> = React.memo(({ value, label, description, isActionable, onClick }) => {
    const displayValue = Math.round(value);
    return (
        <div className={`score-card ${isActionable ? 'actionable' : ''}`} onClick={isActionable ? onClick : undefined}>
            <div className="score-card-header">
                <h4>{label}</h4>
                <div className="info-icon-container">
                    <svg className="info-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>
                    <div className="tooltip">{description}</div>
                </div>
            </div>
            <div className="score-value">{displayValue}%</div>
            <div className="progress-bar">
                <div className="progress-bar-inner" style={{ width: `${displayValue}%` }}></div>
            </div>
            {isActionable && <div className="action-indicator">View Details</div>}
        </div>
    )
});

const scoreDefinitions = {
    en: {
      comparisonQuality: {
        CompletionScore: {
          label: "Completion Score",
          definition: "Evaluates if quotes are complete and include all necessary data fields (e.g., delivery, warranty).",
          calculation: "Checks for missing values, empty cells, or incomplete sections relative to a typical RFQ structure. A high score indicates a thorough, low-risk quote.",
        },
        PrecisionScore: {
          label: "Precision Score",
          definition: "Measures the clarity and specificity of the data, reducing ambiguity.",
          calculation: "Assesses the level of detail in descriptions and specifications. 'Plastic badge, 0.7mm' scores higher than 'Badge'.",
        },
      },
      quotePerformance: {
        PriceCompetitivenessScore: {
          label: "Price Competitiveness",
          definition: "Compares unit and total prices relative to the other suppliers.",
          calculation: "Benchmarks prices for matched items across the quotes. A high score is awarded to the supplier offering the most competitive pricing.",
        },
        DeliveryServiceScore: {
          label: "Delivery & Service",
          definition: "Compares non-price factors like delivery lead times, warranties, and service conditions.",
          calculation: "Evaluates terms related to delivery speed, warranty duration, and included services. A high score indicates superior service and logistical terms.",
        },
      },
    },
    fr: {
      comparisonQuality: {
        CompletionScore: {
          label: "Score de Complétude",
          definition: "Évalue si les devis sont complets et incluent toutes les données nécessaires (ex: livraison, garantie).",
          calculation: "Recherche les valeurs manquantes ou les sections incomplètes. Un score élevé indique un devis approfondi et à faible risque.",
        },
        PrecisionScore: {
          label: "Score de Précision",
          definition: "Mesure la clarté et la spécificité des données, réduisant l'ambiguïté.",
          calculation: "Évalue le niveau de détail dans les descriptions. 'Badge plastique, 0.7mm' obtient un meilleur score que 'Badge'.",
        },
      },
      quotePerformance: {
        PriceCompetitivenessScore: {
          label: "Compétitivité Prix",
          definition: "Compare les prix unitaires et totaux par rapport aux autres fournisseurs.",
          calculation: "Étalonne les prix des articles correspondants entre les devis. Un score élevé est attribué au fournisseur le plus compétitif.",
        },
        DeliveryServiceScore: {
          label: "Livraison & Service",
          definition: "Compare les facteurs non tarifaires comme les délais, les garanties et les conditions de service.",
          calculation: "Évalue les conditions de livraison, la durée de la garantie et les services inclus. Un score élevé indique des conditions de service supérieures.",
        },
      },
    },
  };


const translations = {
    en: {
        appSubtitle: "Intelligently compare procurement quotes to analyze data quality and performance.",
        uploadTitle: "Upload Quotes",
        uploadSubtitle: "Upload at least two Excel or CSV files to begin the comparison.",
        clear: "Clear",
        analyzing: "Analyzing...",
        analyzingQuality: "Analyzing Data Quality...",
        analyzingPerformance: "Analyzing Quote Performance...",
        processingFiles: "Processing local files...",
        contactingAI: "Contacting AI for analysis...",
        receivingAnalysis: "Receiving analysis...",
        reanalyzing: "Re-analyzing...",
        compare: "Compare Quotes",
        resultsPlaceholder: "Your AI-powered comparison report will appear here.",
        comparisonQualityTitle: "Data Quality Check",
        quotePerformanceTitle: "Quote Performance",
        executiveSummary: "Executive Summary",
        recommendation: "Recommendation",
        strengths: "Key Strengths",
        weaknesses: "Key Weaknesses/Risks",
        negotiationPoints: "Negotiation Points",
        downloadPdf: "Download PDF",
        generatingPdf: "Generating PDF...",
        uploadPlaceholder: "Drag & drop files or click to select",
        errorMinTwoFiles: "Please upload at least two quote files to compare.",
        errorAnalyzing: "An error occurred while analyzing the files. Please try again.",
        errorPdf: "Could not generate PDF. Please try again.",
        scoreDefinitionsTitle: "Score Definitions",
        definition: "Definition",
        calculation: "Calculation",
        close: "Close",
        updateScores: "Update Scores",
        item: "Item",
        issue: "Issue",
        correction: "Correction",
        noPointsIdentified: "None identified for this category.",
        correctionNeeded: "Data quality scores are below 100%. For a more accurate analysis, we recommend improving the data. You can proceed now, but results may be less reliable.",
        performanceReady: "Data quality is perfect! You can now proceed to the full analysis.",
        analyzePerformance: "Analyze Performance",
        correctionModalNoIssues: "The score is not 100%, but no specific issues were identified by the AI. Please review the data below and make corrections where needed to improve the score.",
        correctionModalTitle: "Data Correction",
        correctionModalSubtitle: "Standardize the data across quotes to ensure a fair comparison. The AI has highlighted fields that are unclear or inconsistent. Edit the cells below to improve them.",
        itemGroup: "Item Group",
        field: "Field",
        viewBar: "Bar Chart",
        viewTable: "Data Table",
        supplier: "Supplier",
        score: "Score",
        faq: {
            title: "Frequently Asked Questions",
            questions: [
                {
                    q: "How do I use ProcurementCompare AI?",
                    a: "1. **Upload Files**: Drag and drop at least two procurement quotes (Excel or CSV) into the upload area.\n2. **Quality Check**: Click 'Compare Quotes'. The AI will analyze the data quality and provide scores. If scores are below 100%, you can click 'View Details' to correct and standardize the data across files.\n3. **Performance Analysis**: Once data quality is perfect (or you choose to proceed), click 'Analyze Performance'. The AI will generate a full comparison, including performance scores and an executive summary."
                },
                {
                    q: "How is my data processed and stored? Is it secure?",
                    a: "We take your data privacy very seriously. Here's our commitment:\n- **Local Processing**: Files are initially processed directly in your browser. They are not uploaded to our servers at this stage.\n- **AI Analysis**: For analysis, the textual content of your files is sent securely to the Google Gemini API. No files are stored on any server post-analysis.\n- **GDPR Compliant**: The process is designed to be GDPR compliant. We do not store any of your personal or business data. Each session is stateless; once you close the browser tab, all data from the session is gone. Your files remain on your computer."
                },
                {
                    q: "What file formats are supported?",
                    a: "The application supports Microsoft Excel files (.xlsx, .xls) and Comma-Separated Values files (.csv). For best results, ensure your files have a clear header row and consistent formatting."
                }
            ]
        }
    },
    fr: {
        appSubtitle: "Comparez intelligemment les devis d'approvisionnement pour analyser la qualité des données et la performance.",
        uploadTitle: "Télécharger les devis",
        uploadSubtitle: "Téléchargez au moins deux fichiers Excel ou CSV pour commencer la comparaison.",
        clear: "Effacer",
        analyzing: "Analyse en cours...",
        analyzingQuality: "Analyse de la qualité des données...",
        analyzingPerformance: "Analyse de la performance des devis...",
        processingFiles: "Traitement des fichiers locaux...",
        contactingAI: "Contact de l'IA pour analyse...",
        receivingAnalysis: "Réception de l'analyse...",
        reanalyzing: "Ré-analyse en cours...",
        compare: "Comparer les devis",
        resultsPlaceholder: "Votre rapport de comparaison généré par l'IA apparaîtra ici.",
        comparisonQualityTitle: "Vérification de la Qualité des Données",
        quotePerformanceTitle: "Performance des Devis",
        executiveSummary: "Synthèse",
        recommendation: "Recommandation",
        strengths: "Points Clés Forts",
        weaknesses: "Faiblesses/Risques Clés",
        negotiationPoints: "Points de Négociation",
        downloadPdf: "Télécharger PDF",
        generatingPdf: "Génération du PDF...",
        uploadPlaceholder: "Glissez-déposez ou cliquez pour sélectionner",
        errorMinTwoFiles: "Veuillez télécharger au moins deux fichiers de devis pour comparer.",
        errorAnalyzing: "Une erreur s'est produite lors de l'analyse. Veuillez réessayer.",
        errorPdf: "Impossible de générer le PDF. Veuillez réessayer.",
        scoreDefinitionsTitle: "Définitions des Scores",
        definition: "Définition",
        calculation: "Calcul",
        close: "Fermer",
        updateScores: "Mettre à jour les scores",
        item: "Article",
        issue: "Problème",
        correction: "Correction",
        noPointsIdentified: "Aucun point identifié pour cette catégorie.",
        correctionNeeded: "Les scores de qualité des données sont inférieurs à 100%. Pour une analyse plus précise, il est recommandé d'améliorer les données. Vous pouvez continuer, mais les résultats pourraient être moins fiables.",
        performanceReady: "La qualité des données est parfaite ! Vous pouvez maintenant procéder à l'analyse complète.",
        analyzePerformance: "Analyser la Performance",
        correctionModalNoIssues: "Le score n'est pas de 100%, mais aucun problème spécifique n'a été identifié par l'IA. Veuillez examiner les données ci-dessous et apporter les corrections nécessaires pour améliorer le score.",
        correctionModalTitle: "Correction des Données",
        correctionModalSubtitle: "Standardisez les données entre les devis pour garantir une comparaison équitable. L'IA a mis en évidence les champs peu clairs ou incohérents. Modifiez les cellules ci-dessous pour les améliorer.",
        itemGroup: "Groupe d'articles",
        field: "Champ",
        viewBar: "Graphique à Barres",
        viewTable: "Tableau",
        supplier: "Fournisseur",
        score: "Score",
        faq: {
            title: "Foire Aux Questions",
            questions: [
                {
                    q: "Comment utiliser ProcurementCompare IA ?",
                    a: "1. **Télécharger les fichiers**: Glissez-déposez au moins deux devis (Excel ou CSV) dans la zone de téléchargement.\n2. **Contrôle Qualité**: Cliquez sur 'Comparer les devis'. L'IA analysera la qualité des données et fournira des scores. Si les scores sont inférieurs à 100%, cliquez sur 'Voir les détails' pour corriger et standardiser les données.\n3. **Analyse de Performance**: Une fois la qualité des données parfaite (ou si vous choisissez de continuer), cliquez sur 'Analyser la Performance'. L'IA générera une comparaison complète, incluant les scores de performance et une synthèse."
                },
                {
                    q: "Comment mes données sont-elles traitées et stockées ? Est-ce sécurisé ?",
                    a: "Nous prenons la confidentialité de vos données très au sérieux :\n- **Traitement Local**: Vos fichiers sont d'abord traités directement dans votre navigateur. Ils ne sont pas envoyés sur nos serveurs à cette étape.\n- **Analyse IA**: Pour l'analyse, le contenu textuel de vos fichiers est envoyé de manière sécurisée à l'API Google Gemini. Aucun fichier n'est stocké sur un serveur après l'analyse.\n- **Conformité RGPD**: Le processus est conçu pour être conforme au RGPD. Nous ne stockons aucune de vos données personnelles ou professionnelles. Chaque session est sans état ; une fois l'onglet fermé, toutes les données de la session sont effacées. Vos fichiers restent sur votre ordinateur."
                },
                {
                    q: "Quels formats de fichiers sont pris en charge ?",
                    a: "L'application prend en charge les fichiers Microsoft Excel (.xlsx, .xls) et les fichiers CSV (.csv). Pour de meilleurs résultats, assurez-vous que vos fichiers ont une ligne d'en-tête claire et un formatage cohérent."
                }
            ]
        }
    }
};

interface ItemGroup {
    groupId: GroupId;
    items: {
        fileName: string;
        item: ParsedItem;
    }[];
    issues: IssueDetail[];
    commonHeaders: string[];
    bestTitle: string;
}

// FIX: Added a props interface for CorrectionModal to ensure type safety.
interface CorrectionModalProps {
    modalInfo: { scoreKey: keyof ComparisonQualityScores; issues: IssueDetail[] };
    editableData: ParsedData;
    onDataChange: (data: ParsedData) => void;
    onClose: () => void;
    onRecompare: () => void;
    isLoading: boolean;
    scoreDefinitions: (typeof scoreDefinitions)['en'];
    translations: (typeof translations)['en'];
}


const CorrectionModal: React.FC<CorrectionModalProps> = React.memo(({ modalInfo, editableData, onDataChange, onClose, onRecompare, isLoading, scoreDefinitions, translations }) => {
    const { scoreKey, issues } = modalInfo;
    const scoreDef = scoreDefinitions.comparisonQuality[scoreKey];

    const itemGroups = React.useMemo<ItemGroup[]>(() => {
        const groups = new Map<GroupId, ItemGroup>();

        // 1. Group all items by groupId
        for (const [fileName, fileItems] of Object.entries(editableData)) {
            for (const item of fileItems) {
                if (!item.groupId) continue;
                if (!groups.has(item.groupId)) {
                    groups.set(item.groupId, {
                        groupId: item.groupId,
                        items: [],
                        issues: [],
                        commonHeaders: [],
                        bestTitle: `Group ${item.groupId}`
                    });
                }
                groups.get(item.groupId)!.items.push({ fileName, item });
            }
        }
        
        // 2. Add issues to their respective groups
        for (const issue of issues) {
            if (groups.has(issue.groupId)) {
                groups.get(issue.groupId)!.issues.push(issue);
            }
        }

        // 3. Determine common headers and best title for each group
        groups.forEach(group => {
            const allHeaders = new Set<string>();
            let bestTitle = `Item ${group.groupId}`;
            let maxTitleLength = 0;

            group.items.forEach(({ item }) => {
                Object.keys(item).forEach(key => {
                    if (key !== 'id' && key !== 'groupId') {
                        allHeaders.add(key);
                    }
                });
                // Attempt to find a descriptive title
                const potentialTitle = item['Description'] || item['Name'] || item['Product'] || '';
                if (typeof potentialTitle === 'string' && potentialTitle.length > maxTitleLength) {
                    bestTitle = potentialTitle;
                    maxTitleLength = potentialTitle.length;
                }
            });
            group.commonHeaders = Array.from(allHeaders);
            group.bestTitle = bestTitle;
        });

        // 4. Filter for groups that have issues relevant to the current score being viewed
        const relevantGroups = Array.from(groups.values()).filter(g => g.issues.length > 0);

        // If there are no specific issues from the AI, show all groups for manual review
        if(relevantGroups.length === 0) {
            return Array.from(groups.values());
        }

        return relevantGroups;
    }, [editableData, issues]);

    const handleInputChange = (fileName: string, itemId: ItemId, field: string, value: string) => {
        onDataChange({
            ...editableData,
            [fileName]: editableData[fileName].map(item =>
                item.id === itemId ? { ...item, [field]: value } : item
            ),
        });
    };

    const hasIssuesFromAI = issues && issues.length > 0;
    
    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="correction-modal-title">
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="correction-modal-title">{scoreDef.label}: {translations.correctionModalTitle}</h3>
                    <button className="modal-close-btn" onClick={onClose} aria-label={translations.close}>&times;</button>
                </div>
                <div className="modal-body">
                    <p>{hasIssuesFromAI ? translations.correctionModalSubtitle : translations.correctionModalNoIssues}</p>
                    <div className="correction-groups-container">
                        {itemGroups.map(group => {
                            const problematicHeadersFromAI = new Set(group.issues.map(i => i.problematicField).filter(Boolean));
                            
                            return (
                                <div className="correction-group" key={group.groupId}>
                                    <div className="correction-group-header">
                                        <h4>{group.bestTitle}</h4>
                                        <span className="group-id-chip">{translations.itemGroup} #{group.groupId}</span>
                                    </div>

                                    <div className="correction-table-wrapper">
                                        <div className="correction-table" style={{ gridTemplateColumns: `minmax(150px, 1.2fr) repeat(${group.items.length}, minmax(200px, 2fr))` }}>
                                            {/* Table Header */}
                                            <div className="correction-table-header-cell field-header">{translations.field}</div>
                                            {group.items.map(({ fileName }) => (
                                                <div key={fileName} className="correction-table-header-cell file-header">
                                                    <svg className="file-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                                    <span>{fileName}</span>
                                                </div>
                                            ))}

                                            {/* Table Body Rows */}
                                            {group.commonHeaders.map(header => {
                                                const hasAIssueInRow = problematicHeadersFromAI.has(header);

                                                // Client-side discrepancy check if AI provides no issues
                                                let hasClientIssueInRow = false;
                                                if (!hasIssuesFromAI && group.items.length > 1) {
                                                    const values = group.items.map(({ item }) => item[header]);
                                                    const nonEmptyValues = values
                                                        .map(v => (v === null || v === undefined ? '' : String(v).trim()))
                                                        .filter(v => v !== '');

                                                    if (nonEmptyValues.length > 0) {
                                                        const uniqueValues = new Set(nonEmptyValues);
                                                        // A discrepancy exists if there's more than one unique value,
                                                        // or if some items have a value while others don't for this field.
                                                        if (uniqueValues.size > 1 || nonEmptyValues.length < group.items.length) {
                                                            hasClientIssueInRow = true;
                                                        }
                                                    }
                                                }
                                                
                                                const hasIssueInRow = hasAIssueInRow || hasClientIssueInRow;
                                                
                                                return (
                                                    <React.Fragment key={header}>
                                                        <div className={`correction-table-field-cell ${hasIssueInRow ? 'row-has-issue' : ''}`}>{header}</div>
                                                        {group.items.map(({ fileName, item }) => {
                                                            const cellIssue = hasIssuesFromAI ? group.issues.find(
                                                                i => i.itemId === item.id && i.fileName === fileName && i.problematicField === header
                                                            ) : null;
                                                            const isProblematicFromAI = !!cellIssue;
                                                            
                                                            // A cell is problematic if the AI flagged it, or if a client-side discrepancy was found for the row.
                                                            const isProblematic = isProblematicFromAI || hasClientIssueInRow;

                                                            return (
                                                                <div key={`${fileName}-${item.id}-${header}`} className={`correction-table-data-cell ${isProblematic ? 'problematic' : ''}`}>
                                                                    <input
                                                                        type="text"
                                                                        value={item[header] || ''}
                                                                        onChange={(e) => handleInputChange(fileName, item.id, header, e.target.value)}
                                                                    />
                                                                    {isProblematicFromAI && (
                                                                        <div className="info-icon-container">
                                                                            <svg className="info-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                                                                            <div className="tooltip">{cellIssue.explanation}</div>
                                                                        </div>
                                                                    )}
                                                                </div>
                                                            );
                                                        })}
                                                    </React.Fragment>
                                                );
                                            })}
                                        </div>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>
                <div className="modal-footer">
                    <button className="btn btn-secondary" onClick={onClose}>{translations.close}</button>
                    <button className="btn btn-primary" onClick={onRecompare} disabled={isLoading}>
                         {isLoading ? translations.reanalyzing : translations.updateScores}
                    </button>
                </div>
            </div>
        </div>
    );
});

type VisualizationType = 'bar' | 'table';

interface PerformanceVisualizerProps {
    performanceData: QuotePerformance[];
    scoreDefinitions: (typeof scoreDefinitions)['en']['quotePerformance'];
    translations: (typeof translations)['en'];
}

const PerformanceVisualizer: React.FC<PerformanceVisualizerProps> = React.memo(({ performanceData, scoreDefinitions, translations }) => {
    const [view, setView] = React.useState<VisualizationType>('bar');
    const colors = ['#A076F9', '#198754', '#ffc107', '#dc3545', '#6c757d'];

    const renderBarView = () => (
        <div className="performance-by-score-container">
            {Object.keys(scoreDefinitions).map(scoreKey => (
                <div key={scoreKey} className="score-comparison-group">
                    <div className="score-comparison-header">
                        <h4>{scoreDefinitions[scoreKey as keyof QuotePerformanceScores].label}</h4>
                        <div className="info-icon-container">
                            <svg className="info-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>
                            <div className="tooltip">{`${scoreDefinitions[scoreKey as keyof QuotePerformanceScores].definition}\n\n${translations.calculation}: ${scoreDefinitions[scoreKey as keyof QuotePerformanceScores].calculation}`}</div>
                        </div>
                    </div>
                    <div className="supplier-bars">
                        {performanceData.map((perf, index) => (
                            <div key={perf.fileName} className="supplier-bar-item">
                                <span className="supplier-bar-label" title={perf.fileName}>{perf.fileName}</span>
                                <div className="supplier-bar">
                                    <div className="supplier-bar-fill" style={{ width: `${perf.scores[scoreKey as keyof QuotePerformanceScores]}%`, backgroundColor: colors[index % colors.length] }}>
                                        <span>{Math.round(perf.scores[scoreKey as keyof QuotePerformanceScores])}%</span>
                                    </div>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>
            ))}
        </div>
    );

    const renderTableView = () => {
        const scoreKeys = Object.keys(scoreDefinitions);
        const topPerformers: { [key: string]: string } = {};

        scoreKeys.forEach(key => {
            let topFile = '';
            let maxScore = -1;
            performanceData.forEach(perf => {
                if (perf.scores[key as keyof QuotePerformanceScores] > maxScore) {
                    maxScore = perf.scores[key as keyof QuotePerformanceScores];
                    topFile = perf.fileName;
                }
            });
            topPerformers[key] = topFile;
        });

        return (
            <div className="data-table-view">
                <table>
                    <thead>
                        <tr>
                            <th>{translations.supplier}</th>
                            {scoreKeys.map(key => (
                                <th key={key}>{scoreDefinitions[key as keyof QuotePerformanceScores].label}</th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {performanceData.map(perf => (
                            <tr key={perf.fileName}>
                                <td data-label={translations.supplier}>{perf.fileName}</td>
                                {scoreKeys.map(key => (
                                    <td key={key} data-label={scoreDefinitions[key as keyof QuotePerformanceScores].label} className={topPerformers[key] === perf.fileName ? 'top-performer' : ''}>
                                        {Math.round(perf.scores[key as keyof QuotePerformanceScores])}%
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        );
    };
    
    return (
        <div className="performance-visualizer">
            <div className="view-switcher">
                <button className={`view-btn ${view === 'bar' ? 'active' : ''}`} onClick={() => setView('bar')}>{translations.viewBar}</button>
                <button className={`view-btn ${view === 'table' ? 'active' : ''}`} onClick={() => setView('table')}>{translations.viewTable}</button>
            </div>
            <div className="visualization-content">
                {view === 'bar' && renderBarView()}
                {view === 'table' && renderTableView()}
            </div>
        </div>
    );
});

interface FaqModalProps {
    onClose: () => void;
    content: {
        title: string;
        questions: { q: string; a: string; }[];
    };
    translations: { close: string };
}

const FaqModal: React.FC<FaqModalProps> = ({ onClose, content, translations }) => {
    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="faq-modal-title">
            <div className="modal-content faq-modal" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="faq-modal-title">{content.title}</h3>
                    <button className="modal-close-btn" onClick={onClose} aria-label={translations.close}>&times;</button>
                </div>
                <div className="modal-body">
                    <div className="faq-content">
                        {content.questions.map((item, index) => (
                            <details key={index} className="faq-item">
                                <summary className="faq-question">{item.q}</summary>
                                <div className="faq-answer">
                                    {item.a.split('\n').map((line, i) => <p key={i} dangerouslySetInnerHTML={{ __html: line.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>') }} />)}
                                </div>
                            </details>
                        ))}
                    </div>
                </div>
            </div>
        </div>
    );
};

const App = () => {
    const [files, setFiles] = React.useState<File[]>([]);
    const [isLoading, setIsLoading] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string | null>(null);
    const [results, setResults] = React.useState<ComparisonResult | null>(null);
    const [language, setLanguage] = React.useState<'en' | 'fr'>('en');
    const [editableData, setEditableData] = React.useState<ParsedData | null>(null);
    const [modalInfo, setModalInfo] = React.useState<{ scoreKey: keyof ComparisonQualityScores; issues: IssueDetail[] } | null>(null);
    const [stage, setStage] = React.useState<'initial' | 'quality' | 'performance'>('initial');
    const [loadingMessage, setLoadingMessage] = React.useState('');
    const [isFaqOpen, setIsFaqOpen] = React.useState<boolean>(false);

    const t = React.useMemo(() => translations[language], [language]);
    const scoreDefs = React.useMemo(() => scoreDefinitions[language], [language]);

    const handleAddFiles = React.useCallback((newFiles: File[]) => {
        setFiles(prevFiles => {
            const uniqueNewFiles = newFiles.filter(newFile => 
                !prevFiles.some(existingFile => existingFile.name === newFile.name)
            );
            return [...prevFiles, ...uniqueNewFiles];
        });
        setError(null);
    }, []);

    const handleRemoveFile = React.useCallback((fileName: string) => {
        setFiles(prevFiles => prevFiles.filter(file => file.name !== fileName));
    }, []);

    const handleClear = React.useCallback(() => {
        setFiles([]);
        setResults(null);
        setError(null);
        setEditableData(null);
        setStage('initial');
    }, []);
    
    const processFileContent = (file: File): Promise<string> => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsArrayBuffer(file);
            reader.onload = (e) => {
                try {
                    const data = e.target!.result;
                    const workbook = XLSX.read(data, { type: 'array' });
                    let content = '';
                    workbook.SheetNames.forEach((sheetName, index) => {
                        const sheet = workbook.Sheets[sheetName];
                        const csv = XLSX.utils.sheet_to_csv(sheet);
                        if (csv.length > 0) {
                            if (workbook.SheetNames.length > 1) {
                                content += `--- SHEET: ${sheetName} ---\n`;
                            }
                            content += csv;
                            if (index < workbook.SheetNames.length - 1) {
                                content += '\n\n';
                            }
                        }
                    });
                    resolve(content);
                } catch (error) {
                    console.error("Error processing file:", file.name, error);
                    reject(`Error processing file ${file.name}. It might be corrupted or in an unsupported format.`);
                }
            };
            reader.onerror = (error) => reject(error);
        });
    };

    const runComparison = React.useCallback(async (analyzePerformance: boolean, dataToProcess: ParsedData | null = null) => {
        setIsLoading(true);
        setError(null);
        setLoadingMessage(analyzePerformance ? t.analyzingPerformance : t.analyzingQuality);

        const responseSchema = {
          type: Type.OBJECT,
          properties: {
            comparisonQuality: {
              type: Type.OBJECT,
              properties: {
                CompletionScore: { type: Type.NUMBER },
                PrecisionScore: { type: Type.NUMBER },
              },
              required: ['CompletionScore', 'PrecisionScore'],
            },
            detailedIssues: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  issueType: { type: Type.STRING },
                  fileName: { type: Type.STRING },
                  itemId: { type: Type.INTEGER },
                  groupId: { type: Type.INTEGER },
                  explanation: { type: Type.STRING },
                  problematicField: { type: Type.STRING },
                },
                required: ['issueType', 'fileName', 'itemId', 'groupId', 'explanation'],
              },
            },
            parsedDataJSON: {
              type: Type.STRING,
              description: 'A minified JSON string representing the parsed data. The root of this JSON string should be an object where keys are filenames and values are arrays of parsed items.',
            },
            quotePerformance: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  fileName: { type: Type.STRING },
                  scores: {
                    type: Type.OBJECT,
                    properties: {
                      PriceCompetitivenessScore: { type: Type.NUMBER },
                      DeliveryServiceScore: { type: Type.NUMBER },
                    },
                    required: ['PriceCompetitivenessScore', 'DeliveryServiceScore'],
                  },
                },
                required: ['fileName', 'scores'],
              },
            },
            executiveSummary: {
              type: Type.OBJECT,
              properties: {
                recommendation: { type: Type.STRING },
                strengths: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      fileName: { type: Type.STRING },
                      points: { type: Type.ARRAY, items: { type: Type.STRING } },
                    },
                    required: ['fileName', 'points'],
                  },
                },
                weaknesses: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      fileName: { type: Type.STRING },
                      points: { type: Type.ARRAY, items: { type: Type.STRING } },
                    },
                    required: ['fileName', 'points'],
                  },
                },
                negotiationPoints: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      fileName: { type: Type.STRING },
                      points: { type: Type.ARRAY, items: { type: Type.STRING } },
                    },
                    required: ['fileName', 'points'],
                  },
                },
              },
              required: ['recommendation', 'strengths', 'weaknesses', 'negotiationPoints'],
            },
          },
          required: ['comparisonQuality', 'detailedIssues', 'parsedDataJSON'],
        };
        
        try {
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
            
            setLoadingMessage(t.processingFiles);
            const fileContents = dataToProcess ? [] : await Promise.all(
                files.map(async (file) => {
                    const content = await processFileContent(file);
                    return {
                        name: file.name,
                        content: content,
                    };
                })
            );

            const prompt = `Analyze the provided procurement quotes. The content of each file is provided as plain text, typically in CSV format.
            - The user language is: ${language}.
            - The current analysis stage is: ${analyzePerformance ? 'performance' : 'quality'}.
            ${dataToProcess ? '- The user has corrected the data. Re-analyze this updated data.' : ''}
            
            Your primary task is to parse the data, then perform analysis based on the stage.

            Stage 'quality':
            1. Parse each file's content into a structured list of items. Assign a unique numeric 'id' to each item *within that file*.
            2. Critically, you must match corresponding items across all files. An item in one file corresponds to an item in another if they refer to the same real-world product. Use a combination of clues like product descriptions, SKU, part numbers, and even similar pricing to make a confident match. Assign a shared numeric 'groupId' to each set of matched items. If an item is truly unique and has no counterpart in other files, it should still get its own unique 'groupId'. Every single item must have a 'groupId'.
            3. Create a JSON object for the parsed data where keys are filenames and values are the arrays of parsed items from step 1 & 2. Stringify this JSON object and place it in the 'parsedDataJSON' field of your response.
            4. Calculate ComparisonQualityScores (CompletionScore, PrecisionScore) from 0-100.
            5. Identify every specific issue contributing to scores below 100 in 'detailedIssues'. For each issue, provide 'fileName', 'itemId', the corresponding 'groupId', 'issueType', a clear 'explanation', and the 'problematicField' if applicable.

            Stage 'performance':
            1. Use the provided (and corrected) data from the 'Corrected Data' input. DO NOT re-parse the original file contents. The structure will be the same as the one you generate for 'parsedDataJSON'.
            2. Calculate ComparisonQualityScores (should be 100).
            3. Calculate QuotePerformance scores for each file.
            4. Generate the ExecutiveSummary with recommendation, strengths, weaknesses, and negotiation points, with specific points attributed to each file.
            5. The 'parsedDataJSON' field in the output should contain the stringified version of the corrected data you received as input.
            
            You MUST return a single, minified JSON object that strictly conforms to the provided response schema. Do not include any text or explanations outside of the JSON object.`;

            const parts = [
                { text: prompt },
                ...(dataToProcess
                  ? [{ text: `Corrected Data: ${JSON.stringify(dataToProcess)}` }]
                  : fileContents.flatMap(fc => [
                      { text: `--- START OF FILE: ${fc.name} ---` },
                      { text: fc.content },
                      { text: `--- END OF FILE: ${fc.name} ---` }
                    ])),
              ];

            setLoadingMessage(t.contactingAI);
            const responseStream = await ai.models.generateContentStream({
                model: "gemini-2.5-flash",
                contents: parts,
                config: {
                    responseMimeType: 'application/json',
                    responseSchema: responseSchema,
                }
            });

            let accumulatedText = "";
            let firstChunkReceived = false;
            for await (const chunk of responseStream) {
                if (!firstChunkReceived) {
                    firstChunkReceived = true;
                    setLoadingMessage(t.receivingAnalysis);
                }
                accumulatedText += chunk.text;
            }

            // Find the start and end of the JSON object to handle potential extraneous text.
            const firstBracket = accumulatedText.indexOf('{');
            const lastBracket = accumulatedText.lastIndexOf('}');
            if (firstBracket === -1 || lastBracket === -1 || lastBracket < firstBracket) {
                // For debugging, log the problematic response.
                console.error("Invalid AI response:", accumulatedText);
                throw new Error("Could not find a valid JSON object in the AI response.");
            }
            const jsonString = accumulatedText.substring(firstBracket, lastBracket + 1);
            const parsedJson = JSON.parse(jsonString);


            if (typeof parsedJson.parsedDataJSON !== 'string') {
                throw new Error("API response is missing the 'parsedDataJSON' field.");
            }

            const resultData: ComparisonResult = {
                ...parsedJson,
                parsedData: JSON.parse(parsedJson.parsedDataJSON || '{}'),
            };
            delete (resultData as any).parsedDataJSON;

            if (analyzePerformance) {
                setResults(prevResults => ({
                    ...resultData,
                    comparisonQuality: prevResults!.comparisonQuality,
                    detailedIssues: prevResults!.detailedIssues,
                }));
            } else {
                setResults(resultData);
            }
            setEditableData(resultData.parsedData);
            setStage(analyzePerformance ? 'performance' : 'quality');

        } catch (err) {
            console.error(err);
            setError(t.errorAnalyzing);
        } finally {
            setIsLoading(false);
            setLoadingMessage('');
        }
    }, [files, language, t]);
    
    const handleCompare = React.useCallback(() => {
        if (files.length < 2) {
            setError(t.errorMinTwoFiles);
            return;
        }
        runComparison(false);
    }, [files.length, t.errorMinTwoFiles, runComparison]);

    const handleRecompare = React.useCallback(() => {
        if (!editableData) return;
        setModalInfo(null);
        runComparison(false, editableData);
    }, [editableData, runComparison]);

    const handleAnalyzePerformance = React.useCallback(() => {
        if (!editableData) return;
        runComparison(true, editableData);
    }, [editableData, runComparison]);

    const openCorrectionModal = React.useCallback((scoreKey: keyof ComparisonQualityScores) => {
        if (!results) return;
        const relevantIssues = results.detailedIssues.filter(
            issue => issue.issueType === scoreKey
        );
        setModalInfo({ scoreKey, issues: relevantIssues });
    }, [results]);

    const handleDownloadPdf = React.useCallback(async () => {
        const reportElement = document.getElementById('report-content');
        if (!reportElement) return;
    
        setIsLoading(true);
        setLoadingMessage(t.generatingPdf);
        setError(null);
        
        let printContainer: HTMLDivElement | null = null;
    
        try {
            // 1. Create a temporary container for rendering
            printContainer = document.createElement('div');
            printContainer.style.position = 'absolute';
            printContainer.style.left = '-9999px';
            printContainer.style.top = '0';
            printContainer.style.width = '840px'; // Fixed width for consistent A4 rendering
            printContainer.style.padding = '2rem';
            printContainer.style.backgroundColor = 'white';
    
            const contentToPrint = reportElement.cloneNode(true) as HTMLElement;
            contentToPrint.classList.add('pdf-export');
            printContainer.appendChild(contentToPrint);
            document.body.appendChild(printContainer);
            
            // 2. Render the entire report to a single, high-resolution canvas
            const canvas = await html2canvas(printContainer, {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff',
            });
            
            // 3. Prepare PDF document (A4 size)
            const pdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();
            const margin = 10;
            const contentWidth = pdfWidth - (margin * 2);
            const pageHeight = pdfHeight - (margin * 2);
    
            // 4. Calculate scaling factor between canvas and PDF
            const imgWidth = canvas.width;
            const imgHeight = canvas.height;
            const contentHeight = (contentWidth / imgWidth) * imgHeight;
    
            // 5. Slice the large canvas into page-sized chunks and add to the PDF
            let yPositionOnCanvas = 0;
            const totalPages = Math.ceil(contentHeight / pageHeight);

            for (let i = 0; i < totalPages; i++) {
                if (i > 0) {
                    pdf.addPage();
                }

                // Calculate the height of the slice for the current page
                const remainingContentHeight = contentHeight - (i * pageHeight);
                const currentPageHeightInMM = Math.min(pageHeight, remainingContentHeight);
                const sliceHeightOnCanvas = (currentPageHeightInMM / contentHeight) * imgHeight;

                const pageCanvas = document.createElement('canvas');
                pageCanvas.width = imgWidth;
                pageCanvas.height = sliceHeightOnCanvas;
                const pageCanvasContext = pageCanvas.getContext('2d')!;
                
                // Draw the slice from the main canvas onto the temporary page canvas
                pageCanvasContext.drawImage(
                    canvas,
                    0, yPositionOnCanvas, // Source x, y
                    imgWidth, sliceHeightOnCanvas, // Source width, height
                    0, 0, // Destination x, y
                    imgWidth, sliceHeightOnCanvas // Destination width, height
                );
                
                pdf.addImage(pageCanvas.toDataURL('image/png'), 'PNG', margin, margin, contentWidth, currentPageHeightInMM);
    
                yPositionOnCanvas += sliceHeightOnCanvas;
            }
    
            pdf.save("procurement-comparison-report.pdf");
    
        } catch (err) {
            console.error("Failed to generate PDF:", err);
            setError(t.errorPdf);
        } finally {
            // 6. Ensure the temporary container is always removed from the DOM
            if (printContainer) {
                document.body.removeChild(printContainer);
            }
            setIsLoading(false);
            setLoadingMessage('');
        }
    }, [t]);

    const qualityScoresPerfect = results &&
        results.comparisonQuality.CompletionScore === 100 &&
        results.comparisonQuality.PrecisionScore === 100;

    return (
        <div className="app-container">
            {modalInfo && editableData && (
                <CorrectionModal
                    modalInfo={modalInfo}
                    editableData={editableData}
                    onDataChange={setEditableData}
                    onClose={() => setModalInfo(null)}
                    onRecompare={handleRecompare}
                    isLoading={isLoading}
                    scoreDefinitions={scoreDefs}
                    translations={t}
                />
            )}
            {isFaqOpen && (
                <FaqModal
                    onClose={() => setIsFaqOpen(false)}
                    content={t.faq}
                    translations={{ close: t.close }}
                />
            )}
            <header>
                <div className="header-title">
                     <svg className="logo" width="40" height="40" viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
                        <path d="M6,2 C4.89,2 4,2.89 4,4 L4,20 C4,21.11 4.89,22 6,22 L18,22 C19.11,22 20,21.11 20,20 L20,8 L14,2 L6,2 Z" />
                     </svg>
                    <h1>ProcurementCompare <span className="h1-accent">IA</span></h1>
                </div>
                <p className="app-subtitle">{t.appSubtitle}</p>
                <div className="header-actions">
                    <div className="language-switcher">
                        <button className={`lang-btn ${language === 'en' ? 'active' : ''}`} onClick={() => setLanguage('en')}>EN</button>
                        <button className={`lang-btn ${language === 'fr' ? 'active' : ''}`} onClick={() => setLanguage('fr')}>FR</button>
                    </div>
                    <button className="btn-icon" onClick={() => setIsFaqOpen(true)} aria-label={t.faq.title}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v15H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
                    </button>
                </div>
            </header>

            <main className="main-content">
                <div className="panel input-panel">
                    <FileUpload
                        files={files}
                        onAddFiles={handleAddFiles}
                        onRemoveFile={handleRemoveFile}
                        title={t.uploadTitle}
                        subtitle={t.uploadSubtitle}
                        placeholderText={t.uploadPlaceholder}
                    />
                    <div className="button-group">
                         <button className="btn btn-secondary" onClick={handleClear} disabled={files.length === 0}>
                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
                            <span>{t.clear}</span>
                        </button>
                        <button className="btn btn-primary" onClick={handleCompare} disabled={isLoading || files.length < 2}>
                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="20" x2="12" y2="4"></line><line x1="4" y1="12" x2="4" y2="20"></line><line x1="20" y1="12" x2="20" y2="20"></line></svg>
                            <span>{t.compare}</span>
                        </button>
                    </div>
                </div>
                <div className="panel output-panel" id="report-content">
                    {isLoading && !results && (
                        <div className="loader">
                            <div className="spinner"></div>
                            <p>{loadingMessage || t.analyzing}</p>
                        </div>
                    )}
                    {error && <div className="error-message">{error}</div>}
                    {!isLoading && !results && !error && (
                         <div className="output-placeholder">
                            <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 24 24" fill="currentColor">
                                <rect x="4" y="12" width="4" height="8" rx="1.5"></rect>
                                <rect x="10" y="4" width="4" height="16" rx="1.5"></rect>
                                <rect x="16" y="8" width="4" height="12" rx="1.5"></rect>
                            </svg>
                            <h3>{t.resultsPlaceholder}</h3>
                        </div>
                    )}
                    {results && (
                        <div className="results-container">
                            <div className="pdf-section">
                                <h3 className="results-category-title">{t.comparisonQualityTitle}</h3>
                                <div className="scores-grid">
                                    {Object.keys(results.comparisonQuality).map(key => {
                                        const scoreKey = key as keyof ComparisonQualityScores;
                                        const scoreValue = results.comparisonQuality[scoreKey];
                                        const isActionable = scoreValue < 100;
                                        return (
                                            <ScoreCard 
                                                key={scoreKey}
                                                label={scoreDefs.comparisonQuality[scoreKey].label}
                                                value={scoreValue}
                                                description={`${scoreDefs.comparisonQuality[scoreKey].definition}\n\n${t.calculation}: ${scoreDefs.comparisonQuality[scoreKey].calculation}`}
                                                isActionable={isActionable}
                                                onClick={isActionable ? () => openCorrectionModal(scoreKey) : undefined}
                                            />
                                        )
                                    })}
                                </div>
                                {stage === 'quality' && (
                                    <div className={`callout-box ${qualityScoresPerfect ? 'success' : 'warning'}`}>
                                        <p>{qualityScoresPerfect ? t.performanceReady : t.correctionNeeded}</p>
                                        <button className="btn btn-primary" onClick={handleAnalyzePerformance} disabled={isLoading}>
                                            {isLoading ? t.analyzing : t.analyzePerformance}
                                        </button>
                                    </div>
                                )}
                            </div>

                            {results.quotePerformance && stage === 'performance' && (
                                <div className="pdf-section">
                                    <h3 className="results-category-title">{t.quotePerformanceTitle}</h3>
                                    <PerformanceVisualizer 
                                        performanceData={results.quotePerformance}
                                        scoreDefinitions={scoreDefs.quotePerformance}
                                        translations={t}
                                    />
                                </div>
                            )}

                             {results.executiveSummary && stage === 'performance' && (
                                <div className="summary-card pdf-section">
                                    <h3 className="results-category-title">{t.executiveSummary}</h3>
                                    <div className="summary-section">
                                        <h4>
                                            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg>
                                            {t.recommendation}
                                        </h4>
                                        <p>{results.executiveSummary.recommendation}</p>
                                    </div>

                                    <div className="summary-section">
                                        <h4>
                                            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z"></path><path d="m9 12 2 2 4-4"></path></svg>
                                            {t.strengths}
                                        </h4>
                                        <div className="summary-details">
                                            {results.executiveSummary.strengths.length > 0 ? results.executiveSummary.strengths.map(s => (
                                                <div className="summary-supplier-group" key={`strength-${s.fileName}`}>
                                                    <div className="supplier-filename">
                                                        <svg className="file-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                                        <span>{s.fileName}</span>
                                                    </div>
                                                    <ul>{s.points.map((p, i) => <li key={i}>{p}</li>)}</ul>
                                                </div>
                                            )) : <p className="no-points-message">{t.noPointsIdentified}</p>}
                                        </div>
                                    </div>

                                    <div className="summary-section">
                                        <h4>
                                            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                                            {t.weaknesses}
                                        </h4>
                                        <div className="summary-details">
                                            {results.executiveSummary.weaknesses.length > 0 ? results.executiveSummary.weaknesses.map(w => (
                                                <div className="summary-supplier-group" key={`weakness-${w.fileName}`}>
                                                     <div className="supplier-filename">
                                                        <svg className="file-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                                        <span>{w.fileName}</span>
                                                    </div>
                                                    <ul>{w.points.map((p, i) => <li key={i} style={{'--bullet-color': 'var(--error-color)'} as React.CSSProperties}>{p}</li>)}</ul>
                                                </div>
                                            )) : <p className="no-points-message">{t.noPointsIdentified}</p>}
                                        </div>
                                    </div>

                                     <div className="summary-section">
                                        <h4>
                                            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 9a2 2 0 0 1-2 2H6l-4 4V4c0-1.1.9-2 2-2h8a2 2 0 0 1 2 2v5Z"></path><path d="M18 9h2a2 2 0 0 1 2 2v9l-4-4h-2a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2Z"></path></svg>
                                            {t.negotiationPoints}
                                        </h4>
                                        <div className="summary-details">
                                            {results.executiveSummary.negotiationPoints.length > 0 ? results.executiveSummary.negotiationPoints.map(n => (
                                                <div className="summary-supplier-group" key={`nego-${n.fileName}`}>
                                                     <div className="supplier-filename">
                                                        <svg className="file-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                                        <span>{n.fileName}</span>
                                                    </div>
                                                    <ul>{n.points.map((p, i) => <li key={i} style={{'--bullet-color': 'var(--primary-color)'} as React.CSSProperties}>{p}</li>)}</ul>
                                                </div>
                                            )) : <p className="no-points-message">{t.noPointsIdentified}</p>}
                                        </div>
                                    </div>
                                    <button className="btn btn-download" onClick={handleDownloadPdf} disabled={isLoading}>
                                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                                        <span>{isLoading ? t.generatingPdf : t.downloadPdf}</span>
                                    </button>
                                </div>
                             )}

                        </div>
                    )}
                </div>
            </main>
        </div>
    );
};

const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<App />);