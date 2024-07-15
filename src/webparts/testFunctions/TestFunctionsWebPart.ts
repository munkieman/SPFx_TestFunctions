import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestFunctionsWebPart.module.scss';
import * as strings from 'TestFunctionsWebPartStrings';

let libCount=0

export interface ITestFunctionsWebPartProps {
  description: string;
  libraryNamePrev: string;
  libraries: string[];
}

export default class TestFunctionsWebPart extends BaseClientSideWebPart<ITestFunctionsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
    <section class="${styles.testFunctions} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div id="libraries"></div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
    </section>`;

    await this.functionOne(false).then(async () => {
      setTimeout(async () => {
        await this.functionTwo();
      },100)
    })
  }

  private async functionOne(flag:boolean) : Promise<void> {
    alert("function one");
    const libraryName : string[] = ["One", "Two","Three", "Four", "Five"];
    const dcDivisions : string[] = ["aaa","bbb","ccc","ddd","eee"];
    
    for (let x = 0; x < libraryName.length; x++) {
      dcDivisions.forEach(async (site,index)=>{
        this.functionThree(flag,site)
          .then(async (response)=> {
            if(response.length>0){
              await this.functionFour(libraryName[x]);
            }
          })
      });
    }
    return;
  }

  private async functionThree(flag:boolean,site:string) : Promise<any> {
    //alert("function three");
    const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
    const dcTitle = site+"_dc";
    const webDC = `https://${tenant_uri}/sites/${dcTitle}/`; 

    return webDC;
  }

  private async functionFour(library:string) : Promise<void>{
    //alert("function four");
    if(library !== this.properties.libraryNamePrev && library!=="Custom"){
      this.properties.libraryNamePrev = library;
      this.properties.libraries[libCount] = library;
      libCount++;
    }

    return;
  }

  private async functionTwo() : Promise<void> {
    //alert("function two");
    let html = "";
    this.properties.libraries.forEach( (name,index)=>{
      html+=`<h4>${name}</h4>`;
    });

    if(this.domElement.querySelector('#libraries') !== null) {
      this.domElement.querySelector('#libraries')!.innerHTML = html;
    }

    this.functionFive();
    return;
  }

  private functionFive() : void {
    alert("function five");    
  }

  protected onInit(): Promise<void> {
    this.properties.libraries=[];

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
