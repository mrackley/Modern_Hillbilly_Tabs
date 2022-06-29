import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import styles from './ModernHillbillyTabsWebPart.module.scss';
import * as strings from 'ModernHillbillyTabsWebPartStrings';

import * as $ from 'jquery';
import * as jQuery from 'jquery';

export interface IModernHillbillyTabsWebPartProps {
  description: string;
  sectionClass: string;
  webpartClass: string;
  tabData: any[];
}

export default class ModernHillbillyTabsWebPart extends BaseClientSideWebPart<IModernHillbillyTabsWebPartProps> {

  public render(): void {

    require('./AddTabs.js');
    require('./AddTabs.css');

    if (this.displayMode == DisplayMode.Read)
    {
      var tabWebPartID = "";
      var zoneDIV = "";
      
      tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id");       
      zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
      
      var tabsDiv = tabWebPartID + "tabs";
      var contentsDiv = tabWebPartID + "Contents";
      
      this.domElement.innerHTML = "<div data-addui='tabs'><div role='tabs' id='"+tabsDiv+"'></div><div role='contents' id='"+contentsDiv+"'></div></div>";

      var thisTabData = this.properties.tabData;
      for(var x in thisTabData)
      {
        $("#"+tabsDiv).append("<div>"+thisTabData[x].TabLabel+"</div>");
        $("#"+contentsDiv).append($("#"+thisTabData[x].WebPartID));
      }

      //@ts-ignore
      RenderTabs();
      } else {
        this.domElement.innerHTML = `
        <div class="${ styles.modernHillbillyTabs }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
              <div class="${ styles.column }">
                <span class="${ styles.title }">Modern Hillbilly Tabs By Mark Rackley</span>
                <p class="${ styles.subTitle }">Place Web Parts into Tabs.</p>
                <p class="${ styles.description }">To use Modern Hillbilly Tabs: 
                  <ul>
                    <li>Place this web part in the same section of the page as the web parts you would like to put into tabs.</li> 
                    <li>Add the web parts to the section and then edit the properties of this web part.</li>
                    <li>Click on the button to 'Manage Tab Labels' and then specify the labels for each web part using the property control.</li>
                  </ul> 
                  The other two Web Part Properties are used to identify sections/web parts on the screen. Do not change these values unless you know what you are doing.</p>
                <a href="https://github.com/mrackley/Modern_Hillbilly_Tabs" class="${ styles.button }">
                  <span class="${ styles.label }">View Source on GitHub</span>
                </a>
                <a href="https://www.markrackley.net" class="${ styles.button }">
                  <span class="${ styles.label }">View Blog Post</span>
                </a>
              </div>
            </div>
          </div>
        </div>`;
      }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getZones(): Array<[string,string]> {
    const zones = new Array<[string,string]>();

    var tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id");       
    var zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
    var count = 1;
    $(zoneDIV).find("."+this.properties.webpartClass).each(function(){
      var thisWPID = $(this).attr("id");
      if (thisWPID != tabWebPartID)
      {
        const zoneId = $(this).attr("id");
        let zoneName:string = "Web Part " + count;
        count++;
        zones.push([zoneId, zoneName]);
      }
    });

    return zones;
  }  

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('sectionClass', {
                  label: strings.SectionClass,
                  description: "Class identifier for Page Section, don't touch this if you don't know what it means."
                }),
                PropertyPaneTextField('webpartClass', {
                  label: strings.WebPartClass,
                  description: "Class identifier for Web Part, don't touch this if you don't know what it means."
                }),
                PropertyFieldCollectionData("tabData", {
                  key: "tabData",
                  label: strings.TabLabels,
                  panelHeader: "Specify Labels for Tabs",
                  manageBtnLabel: "Manage Tab Labels",
                  value: this.properties.tabData,
                  fields: [
                    {
                      id: "WebPartID",
                      title: "Web Part",
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this.getZones().map((zone:[string,string]) => {
                        return {
                          key: zone["0"],
                          text: zone["1"],
                        };
                      })

                    },
                    {
                      id: "TabLabel",
                      title: "Tab Label",
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
