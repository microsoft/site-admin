export class Template {
    private _template = null;
    get CanvasContent1(): { [key: string]: string } { return { CanvasContent1: this._template }; }

    // Constructor
    constructor(siteId: string, webId: string, listId: string, folderUrl: string) {
        // Set the template
        this._template = CanvasContent
            .replace(/4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb/g, siteId)
            .replace(/8d297e14-676c-4fc3-880f-b4873451ccbb/g, webId)
            .replace(/47e1d508-a230-4984-8c41-2a272a5b255c/g, listId)
            .replace(/\/sites\/Dev\/SiteAssets\/SitePages\/CopilotDataReadiness/g, folderUrl);
    }

    // Sets an image's unique id
    updateImageId(fileName: string, uniqueId: string) {
        // Update the template based on the file
        switch (fileName) {
            case "1004145298":
                this._template = this._template.replace(/data-uniqueid="20eef9f5-daae-4c4b-8e2e-99fcd2f724d9"/g, uniqueId);
                break;
            case "178240801":
                this._template = this._template.replace(/data-uniqueid="f06e61dc-1933-452d-9cda-1a9dfd14c47e"/g, uniqueId);
                break;
            case "1439995666":
                this._template = this._template.replace(/data-uniqueid="d4440c71-3b19-4245-aeff-b5c82a1c1ab3"/g, uniqueId);
                break;
            case "2009954068":
                this._template = this._template.replace(/data-uniqueid="45d11c06-e977-4ecf-8f18-5ef7f9fa357f"/g, uniqueId);
                break;
            case "2403381788":
                this._template = this._template.replace(/data-uniqueid="03bfff8a-355a-45f7-bfeb-5f388541e472"/g, uniqueId);
                break;
            case "2492974796":
                this._template = this._template.replace(/data-uniqueid="1e575195-9b15-423b-90d7-7eb26ba90042"/g, uniqueId);
                break;
            case "2498076700":
                this._template = this._template.replace(/data-uniqueid="c2e27e9e-2c2f-413e-877b-0b3d1514643d"/g, uniqueId);
                break;
            case "3194998129":
                this._template = this._template.replace(/data-uniqueid="c13f26e3-5d11-444b-bfea-d65c0285cfc4"/g, uniqueId);
                break;
            case "3423846637":
                this._template = this._template.replace(/data-uniqueid="ff0c726d-307b-4f7f-a007-a2880713f5e0"/g, uniqueId);
                break;
            case "3608935117":
                this._template = this._template.replace(/data-uniqueid="3548b591-6f2d-48ed-82cb-4511d14d1112"/g, uniqueId);
                break;
            case "3614134656":
                this._template = this._template.replace(/data-uniqueid="9578c9cd-90f1-4de8-8429-0a924708011c"/g, uniqueId);
                break;
            case "3639654909":
                this._template = this._template.replace(/data-uniqueid="c9d0727f-e330-4999-8e8f-f9164d1ccf32"/g, uniqueId);
                break;
            case "3707825808":
                this._template = this._template.replace(/data-uniqueid="f0ff99b6-3ddd-4d1e-8c1a-0a5161135d0a"/g, uniqueId);
                break;
            case "394511437":
                this._template = this._template.replace(/data-uniqueid="b5039e67-e155-41fd-bf46-7423bc33c3cd"/g, uniqueId);
                break;
            case "4193189816":
                this._template = this._template.replace(/data-uniqueid="764ec5e0-430b-4de2-9f4d-cce3738bf310"/g, uniqueId);
                break;
            case "704156399":
                this._template = this._template.replace(/data-uniqueid="9fff543e-e3b6-4015-97e6-f8b368feb186"/g, uniqueId);
                break;
            case "734992370":
                this._template = this._template.replace(/data-uniqueid="4f4d503d-9524-42ae-9fe8-afe52bdfbe95"/g, uniqueId);
                break;
            case "782453716":
                this._template = this._template.replace(/data-uniqueid="c732bee6-491d-4f1e-b4c5-ccfb3012e88f"/g, uniqueId);
                break;
            case "878621426":
                this._template = this._template.replace(/data-uniqueid="bde8a163-4d8c-46f7-802a-c49c8e9ae14b"/g, uniqueId);
                break;
            case "926485291":
                this._template = this._template.replace(/data-uniqueid="48a4e7be-7aa1-4b65-81ea-dcf493f23fb4"/g, uniqueId);
                break;
            case "956444907":
                this._template = this._template.replace(/data-uniqueid="e3a17d56-0de4-4449-854d-160b546c19f0"/g, uniqueId);
                break;
        }
    }
}

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;fed1887c-145d-42fc-a459-ffc27f0cd3fa&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;emphasis&quot;&#58;&#123;&quot;zoneEmphasis&quot;&#58;0&#125;,&quot;id&quot;&#58;&quot;2cb8b5b5-f6dd-47ca-8740-8acc94e95714&quot;,&quot;controlType&quot;&#58;3,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1257,&quot;reservedHeight&quot;&#58;228&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.6"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;2cb8b5b5-f6dd-47ca-8740-8acc94e95714&quot;,&quot;title&quot;&#58;&quot;Banner&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.6&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Copilot Data Readiness&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;isFullWidth&quot;&#58;true,&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|me@gccdatta.onmicrosoft.com&quot;],&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|me@gccdatta.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;me@gccdatta.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;me@gccdatta.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;Gunjan Datta&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;4d99bc03-6948-4eb5-9c38-61db94f9f29b&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;37fbaa23-1356-4bcc-b94e-311d7f5c5250&quot;,&quot;controlType&quot;&#58;4,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">For Copilot to be enabled in the SPO
                environment, the site administrators will need to take action for protecting content from over-sharing.
            </p>
            <ul class="customListStyle" style="list-style-type&#58;disc;">
                <li>
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Cleanup Content</p>
                </li>
                <li>
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Analyze Permissions</p>
                </li>
                <li>
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Apply Site Configurations</p>
                </li>
            </ul>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">This is critical for ensuring AI does
                not expose data to users who do not have a need to know.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Audit Tools</h4>
            <p aria-hidden="true">&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Audit Tools" data-imagenaturalheight="633"
                data-imagenaturalwidth="1821"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/956444907.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="e3a17d56-0de4-4449-854d-160b546c19f0"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="633" data-width="1821"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The <strong>Audit Tools</strong> tab has
                many reports available to the site administrator. We will utilize these reports for completing the
                requested tasks.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;9ca5a388-541d-40e8-8bc9-e188634a2908&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;zoneGroupMetadata&quot;&#58;&#123;&quot;type&quot;&#58;1,&quot;isExpanded&quot;&#58;true,&quot;showDividerLine&quot;&#58;false,&quot;iconAlignment&quot;&#58;&quot;left&quot;,&quot;displayName&quot;&#58;&quot;Content Cleanup&quot;&#125;,&quot;id&quot;&#58;&quot;8e212d92-5a19-4576-96e9-b4f55c60a666&quot;,&quot;controlType&quot;&#58;4,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">More content, more problems. We need to
                ensure the most active/relevant data is referenced by AI, when responding to you.</p>
            <p class="noSpacingAbove noSpacingBelow" aria-hidden="true" data-text-type="noSpacing">&#160;</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Purge Old Content</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Document Retention Report" data-imagenaturalheight="334"
                data-imagenaturalwidth="1014"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/704156399.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="9fff543e-e3b6-4015-97e6-f8b368feb186"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="334" data-width="1014"
                data-widthpercentage="86.44501278772378" data-uploading="0"></div>
            <p>We will be using the <strong>Document Retention</strong> report to find content older than 3 years. Run
                this report, and export the data for review with the site owner(s). Purge data following records
                management policy.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Find Inactive Sites</h4>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Set Search Property" data-imagenaturalheight="209"
                data-imagenaturalwidth="521"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/1439995666.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="d4440c71-3b19-4245-aeff-b5c82a1c1ab3"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="209.40115163147792" data-width="522"
                data-widthpercentage="44.50127877237852" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The first task to complete, is to set
                the custom search property for all sites owned by the grouping. This will allow you to run the report to
                find inactive sites. This may take some time for the search api to refresh.</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Inactive Site Report" data-imagenaturalheight="556"
                data-imagenaturalwidth="994"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/2403381788.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="03bfff8a-355a-45f7-bfeb-5f388541e472"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="554.8812877263581"
                data-width="991.9999999999999" data-widthpercentage="84.56947996589939" data-uploading="0"></div>
            <p>Tagging your sites will allow you to run a report against the search api to find the usage of them.</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sample Inactive Site Report" data-imagenaturalheight="556"
                data-imagenaturalwidth="1659"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/2009954068.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="45d11c06-e977-4ecf-8f18-5ef7f9fa357f"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="556" data-width="1659"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The report will have the last 7/14/30/60/90 views, lifetime views and last modified time. This
                information will allow the site owners determine which sites should be purged and archived.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;4,&quot;zoneId&quot;&#58;&quot;78923fd0-a326-43e1-8899-ad9a645dfb3e&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;zoneGroupMetadata&quot;&#58;&#123;&quot;type&quot;&#58;1,&quot;isExpanded&quot;&#58;true,&quot;showDividerLine&quot;&#58;false,&quot;iconAlignment&quot;&#58;&quot;left&quot;,&quot;displayName&quot;&#58;&quot;Oversharing/Permissions Cleanup&quot;&#125;,&quot;id&quot;&#58;&quot;2c4124fc-c656-468f-95b0-2f7d25f97ad8&quot;,&quot;controlType&quot;&#58;4,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">Oversharing is a major issue with
                content. We have various reports available to help identify and secure the content.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review External Sharing</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Sharing Report" data-imagenaturalheight="368"
                data-imagenaturalwidth="1799"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/394511437.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="b5039e67-e155-41fd-bf46-7423bc33c3cd"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="368" data-width="1799"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The <strong>External Shares</strong>
                report will display content where its search metadata property <strong>ViewableByExternalUsers</strong>
                is set. This will help identify content that is viewable by external users. There are actions that can
                be taken to view/download/delete the content. Export this report and review it with the site owners to
                review the content.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review External Users</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Users Report" data-imagenaturalheight="382"
                data-imagenaturalwidth="1808"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/178240801.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="f06e61dc-1933-452d-9cda-1a9dfd14c47e"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="382" data-width="1808"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The <strong>External Users</strong>
                report will inspect the <strong>User Information List</strong> for any users that have an external
                account. Export this report and review it with the site owners to determine which external users still
                need access to the site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review Sharing Links Site Groups</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sharing Links Report" data-imagenaturalheight="204"
                data-imagenaturalwidth="1811"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/2492974796.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="1e575195-9b15-423b-90d7-7eb26ba90042"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="204" data-width="1811"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">When files are shared, a hidden
                <strong>Sharing Links</strong> site group is created. The <strong>Sharing Links</strong> report will
                show these hidden groups and the associated file with the group. There are actions to view the
                file/group and delete the group. Export this report and review it with the site owners to determine if
                these shared need to be removed.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review EEEU Usage</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="EEEU Report" data-imagenaturalheight="240"
                data-imagenaturalwidth="1797"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/878621426.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="bde8a163-4d8c-46f7-802a-c49c8e9ae14b"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="240" data-width="1797"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The use of <strong>Everyone</strong> and
                <strong>EEEU</strong> (Everyone Except External Users) groups have been used incorrectly. There are
                actions to remove the group from the site, which will remove it from all site groups, or to remove it
                from a group. Export this report and review it with the site owners to determine if everyone within the
                tenant should have access to the content.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review Site Permissions</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3639654909.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="c9d0727f-e330-4999-8e8f-f9164d1ccf32" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="378" data-imagenaturalwidth="1814" data-widthpercentage="100" data-width="1814"
                data-height="378" data-captiontext="Permissions Report"></div>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The <strong>Permissions</strong> report
                will give an overview of all site users/groups and m365 users/groups that have access to your site.
                There are flags for the <strong>Everyone</strong> and <strong>EEEU</strong> accounts being referenced by
                either site/m365 group. It's recommended to reduce the use of “Site Groups” and replace them with
                existing M365 Groups, to help better manage permissions. The <strong>View Users</strong> action will
                display all users from the site/m365 groups. Export this report and review it with the site owners to
                determine if content is being over shared.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Review Unique Permissions</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3423846637.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="ff0c726d-307b-4f7f-a007-a2880713f5e0" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="650" data-imagenaturalwidth="1798" data-widthpercentage="100" data-width="1798"
                data-height="650" data-captiontext="Unique Permissions Report"></div>
            <p>The <strong>Unique Permissions</strong> report may take some time to run, based on the number of
                sub-sites and lists/libraries. This report will contain all items that no longer inherit its
                permissions. There are actions to view the item, file and to restore the permissions to inherit from its
                parent. Export this list and review it with the site owners to determine if they require broken
                inheritance.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;5,&quot;zoneId&quot;&#58;&quot;7ce9199c-c73d-4285-91da-24b709247f06&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;zoneGroupMetadata&quot;&#58;&#123;&quot;type&quot;&#58;1,&quot;isExpanded&quot;&#58;true,&quot;showDividerLine&quot;&#58;false,&quot;iconAlignment&quot;&#58;&quot;left&quot;,&quot;displayName&quot;&#58;&quot;Site Configuration&quot;&#125;,&quot;id&quot;&#58;&quot;52885553-fb1e-4e2e-b4db-0fd15e42cb56&quot;,&quot;controlType&quot;&#58;4,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h4 class="headingSpacingAbove headingSpacingBelow">Management Tab</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/734992370.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="4f4d503d-9524-42ae-9fe8-afe52bdfbe95" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="514" data-imagenaturalwidth="1815" data-widthpercentage="100" data-width="1815"
                data-height="514" data-captiontext="Management Tab"></div>
            <p>The <strong>Management</strong> tab has two properties to help address over sharing. The <strong>Enable
                    Guest Access</strong> option can be used to ensure content is not shared externally. The ability to
                set the default <strong>Sensitivity Label</strong> is also found here.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Features Tab</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3614134656.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="9578c9cd-90f1-4de8-8429-0a924708011c" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="363" data-imagenaturalwidth="1819" data-widthpercentage="100" data-width="1819"
                data-height="363" data-captiontext="Features Tab"></div>
            <p>The <strong>Features</strong> tab has two properties to help address over sharing. The <strong>Disable
                    Company Wide Sharing Links</strong> can be set to reduce the over sharing of content. The
                <strong>Show/Hide Content From Search Results</strong> can be set to hide content from global search. If
                you want to set this option on a sub-site, then you can select the sub-site and set this property in the
                <strong>Sub-Site</strong> tab.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;6,&quot;zoneId&quot;&#58;&quot;871cbf34-192f-4d69-9e43-e7fe5eb9e7c6&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;zoneGroupMetadata&quot;&#58;&#123;&quot;type&quot;&#58;1,&quot;isExpanded&quot;&#58;true,&quot;showDividerLine&quot;&#58;false,&quot;iconAlignment&quot;&#58;&quot;left&quot;,&quot;displayName&quot;&#58;&quot;Sensitivity Label Guidance&quot;&#125;,&quot;id&quot;&#58;&quot;a96453d1-79a9-4a58-9f6a-ca3dc512e082&quot;,&quot;controlType&quot;&#58;4,&quot;addedFromPersistedData&quot;&#58;true,&quot;isFromSectionTemplate&quot;&#58;false,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">Based on content stored within the site,
                we will provide recommendations for Purview sensitivity labelling.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Keyword Search</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Keyword Search" data-imagenaturalheight="387"
                data-imagenaturalwidth="806"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/1004145298.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="20eef9f5-daae-4c4b-8e2e-99fcd2f724d9"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="387" data-width="806"
                data-widthpercentage="68.71270247229326" data-uploading="0"></div>
            <p>The <strong>Search Documents</strong> audit tool report will allow you to find content by keyword. Based
                on content found, you can set the default sensitivity label for the site/library or apply a sensitivity
                label to all files within a library or folder.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Documents Sample Report" data-imagenaturalheight="226"
                data-imagenaturalwidth="1813"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3707825808.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="f0ff99b6-3ddd-4d1e-8c1a-0a5161135d0a"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="226" data-width="1813"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The sample output report will have the ability to view/download the file, as well as delete it. Discuss
                this report with the site owners and apply sensitivity labels per their guidance or to delete the
                content.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Set Site Default Sensitivity Label</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Set Default Site Sensitivity Label"
                data-imagenaturalheight="534" data-imagenaturalwidth="1824"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3194998129.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="c13f26e3-5d11-444b-bfea-d65c0285cfc4"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="534" data-width="1824"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>Under the <strong>Management</strong> tab, you can set the default sensitivity label. You must apply the
                changes in the <strong>Changes</strong> tab to apply it.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Set Library Default Sensitivity Label</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Run Lists/Libraries Report" data-imagenaturalheight="379"
                data-imagenaturalwidth="1821"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/782453716.png"
                data-listid="47e1d508-a230-4984-8c41-2a272a5b255c" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-uniqueid="c732bee6-491d-4f1e-b4c5-ccfb3012e88f"
                data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb" data-height="379" data-width="1821"
                data-widthpercentage="100" data-uploading="0"></div>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/926485291.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="48a4e7be-7aa1-4b65-81ea-dcf493f23fb4" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="283" data-imagenaturalwidth="1808" data-widthpercentage="100" data-width="1808"
                data-height="283" data-captiontext="Select Library"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Run the <strong>Lists/Libraries</strong>
                audit tool report, and select the <strong>Default Label</strong> option for the target library.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/2498076700.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="c2e27e9e-2c2f-413e-877b-0b3d1514643d" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="386" data-imagenaturalwidth="1204" data-widthpercentage="100" data-width="1204"
                data-height="386" data-captiontext="Set Default Sensitivity Label for Library"></div>
            <p>A form will be displayed for you to select the default sensitivity label.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Set Folder/Files Sensitivity Label</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/3608935117.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="3548b591-6f2d-48ed-82cb-4511d14d1112" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="313" data-imagenaturalwidth="1984" data-widthpercentage="100" data-width="1984"
                data-height="313" data-captiontext="Select Label Files"></div>
            <p>Similar to setting the default label for a library, you have another option to apply sensitivity labels
                to the files in the library. The default label applied doesn't update the content within the library, it
                will only apply to new or updates to existing content. Clicking on the <strong>Label Files</strong>
                option, will display a form to apply sensitivity labels.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Dev/SiteAssets/SitePages/CopilotDataReadiness/4193189816.png" data-uploading="0"
                data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="764ec5e0-430b-4de2-9f4d-cce3738bf310" data-webid="8d297e14-676c-4fc3-880f-b4873451ccbb"
                data-siteid="4d5a1bbe-e47a-43b4-b3a1-3ff33d0c2fcb" data-listid="47e1d508-a230-4984-8c41-2a272a5b255c"
                data-imagenaturalheight="811" data-imagenaturalwidth="1189" data-widthpercentage="100" data-width="1189"
                data-height="811" data-captiontext="Apply Sensitivity Labels"></div>
            <p>By default, files that haven't been tagged will apply the selected label. There is an option to override
                existing labels applied to files, but this may not work based on the current label. This follows rules
                set by Purview and will show the error message in the output report. The ability to select a folder (at
                the root level only) to target is also available, instead of the entire library.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;false&#125;&#125;&#125;">
    </div>
</div>`;