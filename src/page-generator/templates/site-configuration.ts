import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;93768064-52cb-4054-8860-b37f6eb03306&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;e4371c8b-a48c-46f5-88b5-961f3b0d9418&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1629,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;e4371c8b-a48c-46f5-88b5-961f3b0d9418&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&quot;imageSource&quot;&#58;&quot;/_layouts/15/images/sleektemplateimagetile.jpg&quot;&#125;,&quot;links&quot;&#58;&#123;&#125;,&quot;customMetadata&quot;&#58;&#123;&quot;imageSource&quot;&#58;&#123;&#125;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;2,&quot;title&quot;&#58;&quot;SAT - Site Configuration&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - Site Configuration&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""><img data-sp-prop-name="imageSource"
                    src="/_layouts/15/images/sleektemplateimagetile.jpg" /></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;046f5086-68e7-49f5-9202-8117c8f8fdb4&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;cb3f5ab9-bb55-40c3-95be-62e97dd1b0b4&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Site Configuration will need to be
                reviewed by the Site Collection Administrator. Site owners will not have access to these tabs. This
                feature is available in the Full Version of SAT. The audit only setting will not display these tabs.</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The settings will show the current
                configuration of the site. Settings may be disabled by the tenant administrator when configuring this
                webpart.</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><i><strong>[Placeholder to add
                        governance/policies and direction of configuration settings to review]</strong></i></p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;046f5086-68e7-49f5-9202-8117c8f8fdb4&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;f8c010b8-d27e-4baa-b7cd-5341eb1ae2fe&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <p aria-hidden="true">&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Management Tab" data-imagenaturalheight="613"
                data-imagenaturalwidth="1457"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATSiteConfiguration/3007540558.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="473fa072-6544-40e0-84bd-20eec69b23a0"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="613" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p aria-hidden="true">&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Management tab will allow you to
                configure the following settings&#58;</p>
            <ul class="customListStyle" style="list-style-type&#58;disc;">
                <li data-list-item-id="efa1c34f700e99187a684aeb494453386">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Enable Custom
                            Scripts</strong> - Enabling this setting will allow for custom code to be injected into the
                        pages. The classic webparts (Script Editor and Content Editor) will be made available. If set,
                        this will be reset after 24 hours per <a
                            href="https&#58;//learn.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script">documentation</a>.
                    </p>
                </li>
                <li data-list-item-id="ee1718752a791113b48a7e8e27fb6fe1b">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Lock State</strong> -
                        Ability to make a site read-only or prevent any access to it. If “No Access” is selected, then
                        the site will no longer load to any user, including this tool. A helpdesk ticket will need to be
                        created and a powershell script must be run against the site in order to make it read-only or
                        unlocked.</p>
                </li>
                <li data-list-item-id="e0974f8e37a6ba995ecba6ad9fd8e8fa2">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Enable App Catalog
                        </strong>- Setting this option will enable the app catalog for the site.</p>
                </li>
                <li data-list-item-id="eb61a76777aa67aaa5601f308dddb1069">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Increase
                            Storage</strong> - Enabling this option will send a request to update the storage for this
                        site.</p>
                </li>
                <li data-list-item-id="e6100143b9c07b8188576499650e38e6c">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Enable Guest Access
                        </strong>- This setting will enable/disable guest access to the site.</p>
                </li>
                <li data-list-item-id="e7b84e563e60fc639c9c57f9e54810e05">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Sensitivity Label
                        </strong>- Sets the default sensitivity label for the site collection.</p>
                </li>
            </ul>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;046f5086-68e7-49f5-9202-8117c8f8fdb4&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;3&#125;,&quot;id&quot;&#58;&quot;09616416-6ec5-42e6-9dd0-794e3c477cf0&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <p aria-hidden="true">&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Features Tab" data-imagenaturalheight="603"
                data-imagenaturalwidth="1458"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATSiteConfiguration/459558559.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="9feef896-4f41-4d1c-baaf-ef6ff031d839"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="603" data-width="1458"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Feature tab will allow you to
                configure the following settings&#58;</p>
            <ul class="customListStyle" style="list-style-type&#58;disc;">
                <li data-list-item-id="e94ec87b79133d2f9e1756ef4e8c6c662">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Client Side Assets
                            (Flow3)</strong> - Associated with the site collection app catalog. This will set the
                        ExemptFromBlockDownloadOfNonViewableFiles property of the hidden ClientSideAssets library
                        associated with the app catalog. This setting will bypass the Flow3 download exception so SPFx
                        apps will work over Flow3.</p>
                </li>
                <li data-list-item-id="e7fc2eed04bdcce0027a2074447e37a51">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Comments on Site Pages
                            Disabled</strong> - Configuration property to disable comments on site pages.</p>
                </li>
                <li data-list-item-id="e7aac500efa5ed7ec329976642acaba5d">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Disable Company Wide
                            Sharing Links</strong> - Limits the sharing configuration to not allow the option to create
                        a link company wide.</p>
                </li>
                <li data-list-item-id="ede3e0cb798c35141031714522c32ba3c">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Restrict Content
                            Discovery</strong> - Enables/Disables RCD (Restrict Content Discovery) on the site. When
                        enabled, it will remove this site from global search and ability to be read by Copilot.</p>
                </li>
                <li data-list-item-id="e6a4180ebc70bb49c795fcecfbf849bdc">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Remove Site From Search
                            Results</strong> - Removes the site collection from the search index. This will remove the
                        site from global/local search and ability to be read by Copilot.</p>
                </li>
                <li data-list-item-id="e31d4800b649f5f6c38bba32396910b5d">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Social Bar on Site Pages
                            Disabled </strong>- Configuration property to disable the social bar on site pages.</p>
                </li>
                <li data-list-item-id="e5cadda919fc391d13ae3da400e9d7bd7">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>Prevent Apps to
                            Sync</strong> - Configuration property to disable the ability to sync the site with
                        applications.</p>
                </li>
            </ul>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
export const SiteConfiguration = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/site-admin/SiteAssets/SitePages/SATSiteConfiguration", {
    "3007540558": "473fa072-6544-40e0-84bd-20eec69b23a0",
    "459558559": "9feef896-4f41-4d1c-baaf-ef6ff031d839"
});