import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { PropertyFieldCustomList, CustomListFieldType } from 'sp-client-custom-fields/lib/PropertyFieldCustomList';
import { App } from 'botframework-webchat';
import { DirectLine, Message } from 'botframework-directlinejs';
require('../../../node_modules/BotFramework-WebChat/botchat.css');
import styles from './TheHubBot.module.scss';
import * as strings from 'theHubBotStrings';
import { ITheHubBotWebPartProps } from './ITheHubBotWebPartProps';
import { IBotItem } from './components/IBotItem';
import { IBotItems } from './components/IBotItems';

export default class TheHubBotWebPart extends BaseClientSideWebPart<ITheHubBotWebPartProps> {

  private wpInstanceId: string;

  public constructor(context?: IWebPartContext) {
    super();

    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }
  

  public render(): void {
    this.domElement.innerHTML = `<div id="${this.context.instanceId}" class="${styles.thehubbot}"></div>`;
    this.wpInstanceId = this.context.instanceId;

    var user = { id: "userid", name: "unknown" };
    var bot = { id: "userid", name: "unknown" };

    // Get userprofile from SharePoint REST endpoint
    /*
    var req = new XMLHttpRequest();
    req.open("GET", "/_api/SP.UserProfiles.PeopleManager/GetMyProperties", false);
    req.setRequestHeader("Accept", "application/json");
    req.send();
    
    if (req.status == 200) {
      var result = JSON.parse(req.responseText);
      user.id = result.Email;
      user.name = result.DisplayName;
    }
    */

    // Initialize DirectLine connection
    var botConnection = new DirectLine({
      secret: this.properties.mainBotId
    });

    // Initialize the BotChat.App with basic config data and the wrapper element
    App({
      user: user,
      bot: bot,
      botConnection: botConnection
    }, document.getElementById(this.wpInstanceId));

    // Call the bot backchannel to give it user information
    botConnection
      .postActivity({ type: "event", name: "initialize", value: user.name, from: user })
      .subscribe(id => console.log("success initializing"));

      botConnection.activity$
      .filter(activity => activity.type == "message" && activity.from.id == "HollisHomeBot")
      .subscribe(a => {
        debugger;
      });
    // Listen for events on the backchannel
    //for (var item in this.properties.botList) {
      if (this.properties.botList != null && this.properties.botList.length > 0) {
      this.properties.botList.forEach((item, index) => {
    
      var botItem:IBotItem = this.properties.botList[index];
      botItem.InstanceId = this.context.instanceId;
      
      
      botConnection.activity$
      .filter(activity => activity.type == "message" && activity.from.id == "HollisHomeBot" && activity.text.indexOf('@' + botItem.Title) > -1)
      .subscribe(a => {
          debugger;
        var act: any = a;
        var messageText = act.text;
        var botConnection = new DirectLine({
          secret: botItem.DirectLineId
        });
        
        var user = { id: "userid", name: "unknown" };
        var bot = { id: "userid", name: "unknown" };

        document.getElementById(botItem.InstanceId).innerHTML = '<div>loading</div>';
        // Initialize the BotChat.App with basic config data and the wrapper element
        App({
          user: user,
          bot: bot,
          botConnection: botConnection
        }, document.getElementById(botItem.InstanceId));
        console.log('Need to change the direct line bot');

        // Call the bot backchannel to give it user information
        botConnection
          .postActivity({ type: "event", name: "initializeBot", value: messageText, from: user })
          .subscribe(id => console.log("success initializing bot"));
        }
      );
    });

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
                }),
                PropertyPaneTextField('mainBotId', {
                  label: 'Main Bot ID'
                }),
                PropertyFieldCustomList('botList', {
                  label: 'Bots',
                  value: this.properties.botList,
                  headerText: 'List of bots that can be queried',
                  fields: [
                    { id: 'Title', title: 'Title', required: true, type: CustomListFieldType.string },
                    { id: 'DirectLineId', title: 'DirectLineId', required: true, type: CustomListFieldType.string }
                  ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  key: "botListField"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
