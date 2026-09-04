import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;bd9e2e5b-43be-45cf-8bfe-01856b570269&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1469,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;4,&quot;title&quot;&#58;&quot;SAT - Audit Reports&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - Audit Reports&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1,&quot;imageSource&quot;&#58;&quot;&quot;,&quot;altText&quot;&#58;&quot;&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;54fe42bc-9ec4-446f-a389-3b59f45cc95f&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;ba82666e-d7ac-4097-ac1a-ae8c0d392a21&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Audit Tools tab will display the
                available reports for auditing your site collection and sub-webs. The List/Libraries tab will also have
                an option to run reports directly against them. All reports will have the option to export the results
                to a CSV file for review with the admins/owners of the site.</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>[Placeholder to add what reports
                    to run]</strong></p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;54fe42bc-9ec4-446f-a389-3b59f45cc95f&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;86e72020-10f3-43a5-ade1-cc0c657f8e55&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Audit Tools" data-imagenaturalheight="686"
                data-imagenaturalwidth="1452"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1539840079.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="5f225b4c-199a-4831-bd09-c033216bacc8"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="686" data-width="1452"
                data-widthpercentage="100" data-uploading="0"></div>
            <h2 class="headingSpacingAbove headingSpacingBelow">Large Webs/Lists/Libraries</h2>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Target Web will default to search
                all sub-webs for the site. You are able to select a single target web to run the reports against to
                better handle site collections with a large number of sub-webs. Additionally, the Skip Large Lists will
                default to analyzing all lists in the site. This will allow you to run the report against all webs, then
                each large list individually.</p>
            <h2 class="headingSpacingAbove headingSpacingBelow">Large M365 Groups</h2>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The webpart can be configured with a
                list of M365 groups that contain a large number of users. These custom groups will be incorporated with
                various reports that help identify if content is being overshared or not.</p>
            <h2 class="headingSpacingAbove headingSpacingBelow">Available Reports</h2>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsDocRetention.aspx">Data Loss Prevention
                    (DLP)</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The DLP report will help identify
                sensitive content in the site. This requires DLP to be configured in the environment.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsExtShares.aspx">Document Retention</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Document Retention report will help
                identify old content to be removed or archived as part of clean-up of stale data.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsExtUsers.aspx">External Shares</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The External Shares report will help
                identify content that is viewable externally. This report requires the site to be in the search index.
            </p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsExtUsers.aspx">External Users</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The External Users report will help
                identify guest accounts with access to the site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsPermissions.aspx">Permissions</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Permissions report will help
                identify oversharing of content in the site by displaying all role assignments currently in use by the
                site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSearchAgents.aspx">Search Agents</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Agents report will help
                identify any agent files that may exist in a site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSearchDocs.aspx">Search Documents</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Documents report will help
                identify sensitive content by keyword or regex pattern. The keyword search option will require the site
                to be in the search index.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSearchEEEU.aspx">Search Everyone Except
                    External Users (EEEU)</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search EEEU report will help
                identify oversharing by the EEEU, Everyone group and large M365 groups.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSearchProperty.aspx">Search Property</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Property report allows sites
                to be tagged by a configurable set of values the site admin/owner can use for tagging their sites. This
                will help with identifying sites that are orphaned or inactive. This report requires the sites to be in
                the search index.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSearchUsers.aspx">Search Users</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Users report will allow the
                site admin/owner to find users by text or account. This will help with offboarding users to ensure they
                are removed from a site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSensitivityLabels.aspx">Sensitivity Labels</a>
            </h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Sensitivity Labels report will help
                identify files that do not have a label assigned to them. The report will also allow you to find files
                that have been labelled.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsSharingLinks.aspx">Sharing Links</a></h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Sharing Links report will help
                identify files that have been shared.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow"><a
                    href="/sites/demo/site-admin/sitepages/SATAuditReportsUniquePermissions.aspx">Unique Permissions</a>
            </h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Unique Permissions report will help
                identify content that has unique permissions.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
export const AuditReports = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/SiteAssets/SitePages/SATAuditReports", {
        "1539840079": "5f225b4c-199a-4831-bd09-c033216bacc8"
    });