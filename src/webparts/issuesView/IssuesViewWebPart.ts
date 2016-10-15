import {BaseClientSideWebPart, IPropertyPaneSettings, IWebPartContext, PropertyPaneTextField} from '@microsoft/sp-client-preview';
import {PropertyFieldPeoplePicker} from 'sp-client-custom-fields/lib/PropertyFieldPeoplePicker';
import {IPropertyFieldPeople} from 'sp-client-custom-fields/lib/PropertyFieldPeoplePicker';

import styles from './IssuesView.module.scss';
import * as strings from 'issuesViewStrings';
import {IIssuesViewWebPartProps} from './IIssuesViewWebPartProps';

import importableModuleLoader from '@microsoft/sp-module-loader';
import Functions from './Functions'
import MockHttpClient from './MockHttpClient';
import {EnvironmentType} from '@microsoft/sp-client-base';

import {IDisplayItem} from './interfaces/IDisplayItem'
import {IDisplayItems} from './interfaces/IDisplayItem'

require("jquery");
import * as $ from 'jquery';
var ChartJS = require('chartjs');
var pnp = require('sp-pnp-js');

export interface IIssueItems {
    value : IIssueItem[];
}

export interface IIssueItem {
    Title : string;
    Assigned_x0020_To : any;
    Issue_x0020_Status : string;
    Percent_x0020_Complete : number;
    DaysAllocated : number;
}

export default class IssuesViewWebPart extends BaseClientSideWebPart < IIssuesViewWebPartProps > {

    private bgColors : string[] = [
        'rgba(255,99,132, 1.0)',
        'rgba(54, 162, 235, 1.0)',
        'rgba(255, 206, 86, 1.0)',
        'rgba(75, 192, 192, 1.0)',
        'rgba(153, 102, 255, 1.0)',
        'rgba(255, 159, 64, 1.0)'
    ];

    public constructor(context : IWebPartContext) {
        super(context);
        this.onPropertyChange = this
            .onPropertyChange
            .bind(this);
        this.onPropertyChange = this
            .onPropertyChange
            .bind(this);
        this.rendered = this
            .rendered
            .bind(this);
        this.onLoaded = this
            .onLoaded
            .bind(this);

        pnp.setup({
            headers: {
                'Accept': 'application/json;odata=nometadata'
            }
        });
    }

    private _getMockListData() : Promise < IIssueItems > {
        return MockHttpClient
            .get(this.context.pageContext.web.absoluteUrl)
            .then((data : IIssueItem[]) => {
                var listData : IIssueItems = {
                    value: data
                };
                return listData;
            })as Promise < IIssueItems >;
    }

    private _getIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        return pnp
            .sp
            .web
            .lists
            .getByTitle(listName)
            .items
            .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
            .expand('Assigned_x0020_To/Title')
            .filter("Assigned_x0020_To/Title eq '" + personName + "'")
            .get();
    }

    private _getStartedIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        return pnp
            .sp
            .web
            .lists
            .getByTitle(listName)
            .items
            .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
            .expand('Assigned_x0020_To/Title')
            .filter("Assigned_x0020_To/Title eq '" + personName + "' and Issue_x0020_Status eq 'In Progress'")
            .get();
    }

    private _getCompletedIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        return pnp
            .sp
            .web
            .lists
            .getByTitle(listName)
            .items
            .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
            .expand('Assigned_x0020_To/Title')
            .filter("Assigned_x0020_To/Title eq '" + personName + "' and Issue_x0020_Status eq 'Completed'")
            .get();
    }

    private _getIssuesListData(siteUrl : string, personName : string, listName : string) : Promise < IIssueItems > {
        return this
            ._getIssuesDataUsingPnP(siteUrl, personName, listName)
            .then((data : IIssueItem[]) => {
                var listData : IIssueItems = {
                    value: data
                };
                return listData;
            })as Promise < IIssueItems >;
    }

    private _getStartedIssuesListData(siteUrl : string, personName : string, listName : string) : Promise < IIssueItems > {
        return this
            ._getStartedIssuesDataUsingPnP(siteUrl, personName, listName)
            .then((data : IIssueItem[]) => {
                var listData : IIssueItems = {
                    value: data
                };
                return listData;
            })as Promise < IIssueItems >;
    }

    private _getCompletedIssuesListData(siteUrl : string, personName : string, listName : string) : Promise < IIssueItems > {
        return this
            ._getCompletedIssuesDataUsingPnP(siteUrl, personName, listName)
            .then((data : IIssueItem[]) => {
                var listData : IIssueItems = {
                    value: data
                };
                return listData;
            })as Promise < IIssueItems >;
    }

    private _renderIssuesAsync() : void {
        // Local environment
        if(this.context.environment.type === EnvironmentType.Local) {

            var peopleList : IPropertyFieldPeople[] = [];
            if (this.properties && this.properties.people) {
                peopleList = this.properties.people;
                console.log(peopleList[0]['fullName']);
                this
                    ._getMockListData()
                    .then((response) => {
                        const displayItems : IDisplayItem[] = [];
                        for (let i : number = 0; i < response.value.length; i++) {
                            const issueItem : IIssueItem = response.value[i];
                            displayItems.push({Label: issueItem.Title, Data: issueItem.Percent_x0020_Complete, BgColor: this.bgColors[i], HoverBgColor: this.bgColors[i]});
                        }
                        Functions.renderAllIssues(displayItems);
                    });
            }

        } else {
            var peopleList : IPropertyFieldPeople[] = [];
            if (this.properties && this.properties.people) {
                peopleList = this.properties.people;
                console.log(peopleList);

                this._getIssuesListData(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                    const displayItems : IDisplayItem[] = [];
                    for (let i : number = 0; i < response.value.length; i++) {
                        const issueItem : IIssueItem = response.value[i];
                        displayItems.push({Label: issueItem.Title, Data: issueItem.DaysAllocated, BgColor: this.bgColors[i], HoverBgColor: this.bgColors[i]});
                    }
                    Functions.renderAllIssues(displayItems);
                });

                this._getStartedIssuesListData(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                    const displayItems : IDisplayItem[] = [];
                    for (let i : number = 0; i < response.value.length; i++) {
                        const issueItem : IIssueItem = response.value[i];
                        displayItems.push({
                            Label: issueItem.Title,
                            Data: issueItem.Percent_x0020_Complete * 100,
                            BgColor: this.bgColors[i].replace("1.0", "0.2"),
                            HoverBgColor: this.bgColors[i]
                        });
                    }
                    Functions.renderIssuesStarted(displayItems);
                });

                this._getCompletedIssuesListData(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                    let html : string = '';
                    for (let i : number = 0; i < response.value.length; i++) {
                        html += `<li class="ms-font-m">${response.value[i].Title}</li>`;
                    }
                    const completedIssuesContainer : Element = this.domElement.querySelector('#completedIssuesContainer');
                    completedIssuesContainer.innerHTML = html;
                });
            }

        }

    }

    public render() : void {
        var peopleList: IPropertyFieldPeople[] = [];
        var selectedPerson: string = '';
        var selectedPersonImg: string = '';
        var jobTitle: string = '';
        if (this.properties && this.properties.people) {
            peopleList = this.properties.people;
            selectedPerson = peopleList[0]['fullName'];
            selectedPersonImg = peopleList[0]['imageUrl'];
            jobTitle = peopleList[0]['jobTitle'];
        }
        this.domElement.innerHTML = `
        <div class="ms-Grid">
            <p class="ms-font-xxl ${styles.header}">Issue status for</p>
            <div class="ms-Persona">
                <div class="ms-Persona-imageArea">
                    <img class="ms-Persona-image" src="${selectedPersonImg}">
                </div>
                <div class="ms-Persona-details">
                    <div class="ms-Persona-primaryText">${selectedPerson}</div>
                    <div class="ms-Persona-secondaryText">${jobTitle}</div>
                </div>
            </div>
                <div class="ms-Grid-row">
                    <div class="ms-Grid-col ms-u-sm6 ms-u-md4 ms-u-lg12">
                        <p class="ms-font-l ${styles.header}"><i class="ms-Icon ms-Icon--Error" aria-hidden="true"></i> All issues and days allocated</p>
                        <div id="allIssuesChart">
                            <canvas></canvas>
                        </div>
                    </div>
                    <div class="ms-Grid-col ms-u-sm6 ms-u-md8 ms-u-lg12">
                        <p class="ms-font-l ${styles.header}"><i class="ms-Icon ms-Icon--WorkFlow" aria-hidden="true"></i> Issues started</p>
                        <div id="startedIssuesChart">
                            <canvas></canvas>
                        </div>
                    </div>
                    <div class="ms-Grid-col ms-u-sm6 ms-u-md8 ms-u-lg12">
                        <p class="ms-font-l ${styles.header}"><i class="ms-Icon ms-Icon--CompletedSolid" aria-hidden="true"></i> Issues Completed</p>
                        <ul id="completedIssuesContainer">
                        </div>
                    </div>
                </div>
        </div>
`;

        this._renderIssuesAsync();

    }

    private onLoaded() : void {}

    private rendered(carousel : any) {}

    protected get propertyPaneSettings() : IPropertyPaneSettings {
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
                                PropertyFieldPeoplePicker('people', {
                                    label: "Select a user",
                                    initialData: this.properties.people,
                                    allowDuplicate: false,
                                    onPropertyChange: this.onPropertyChange,
                                    context: this.context
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
