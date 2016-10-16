import {BaseClientSideWebPart, IPropertyPaneSettings, IWebPartContext, PropertyPaneTextField, PropertyPaneToggle} from '@microsoft/sp-client-preview';
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

import { IIssuePredictionItem } from './interfaces/IIssuePredictionItem'
import { IIssuePredictionItems } from './interfaces/IIssuePredictionItem'

require("jquery");
require('jqueryui');
import * as $ from 'jquery';
var ChartJS = require('chartjs');
var pnp = require('sp-pnp-js');
var swal = require('swal');

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
        'rgba(255, 159, 64, 1.0)',
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

        //importableModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css');
        importableModuleLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
        importableModuleLoader.loadCss('https://cdn.jsdelivr.net/sweetalert2/5.3.1/sweetalert2.min.css');
        importableModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');
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

    private _getIssuesDataForPredUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssuePredictionItem[] > {

        if(this.properties.showForTeam){
            return pnp
            .sp
            .web
            .lists
            .getByTitle(listName)
            .items
            .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated', 'DaysUsed')
            .expand('Assigned_x0020_To/Title')
            .filter("Assigned_x0020_To/Title eq '" + personName + "'")
            .get();
        }
        else{
            return pnp
            .sp
            .web
            .lists
            .getByTitle(listName)
            .items
            .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated', 'DaysUsed')
            .expand('Assigned_x0020_To/Title')
            .get();
        }


    }

    private _getIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        if(this.properties.showForTeam){
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
        else {
            return pnp
                .sp
                .web
                .lists
                .getByTitle(listName)
                .items
                .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
                .expand('Assigned_x0020_To/Title')
                .get();
            }
    }

    private _getStartedIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        if(this.properties.showForTeam){
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
        else {
            return pnp
                .sp
                .web
                .lists
                .getByTitle(listName)
                .items
                .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
                .expand('Assigned_x0020_To/Title')
                .filter("Issue_x0020_Status eq 'In Progress'")
                .get();
        }
    }

    private _getCompletedIssuesDataUsingPnP(siteUrl : string, personName : string, listName : string) : Promise < IIssueItem[] > {
        if(this.properties.showForTeam){
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
        else {
            return pnp
                .sp
                .web
                .lists
                .getByTitle(listName)
                .items
                .select('Title', 'Id', 'Issue_x0020_Status', 'Percent_x0020_Complete', 'Assigned_x0020_To/Title', 'DaysAllocated')
                .expand('Assigned_x0020_To/Title')
                .filter("Issue_x0020_Status eq 'Completed'")
                .get();
        }
    }

     private _getIssuesListDataForPred(siteUrl : string, personName : string, listName : string) : Promise < IIssuePredictionItems > {
        return this
            ._getIssuesDataForPredUsingPnP(siteUrl, personName, listName)
            .then((data : IIssuePredictionItem[]) => {
                var listData : IIssuePredictionItems = {
                    value: data
                };
                return listData;
            })as Promise < IIssuePredictionItems >;
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
            if(this.properties.showForTeam){
                var peopleList : IPropertyFieldPeople[] = [];
                if (this.properties && this.properties.people) {
                    peopleList = this.properties.people;

                    this.renderAllIssues();

                    //this.renderStartedIssues();

                    this._getCompletedIssuesListData(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                        let html : string = '';
                        for (let i : number = 0; i < response.value.length; i++) {
                            html += `<li class="ms-font-m ms-ListItem">${response.value[i].Title}</li>`;
                        }
                        const completedIssuesContainer : Element = this.domElement.querySelector('#completedIssuesContainer');
                        completedIssuesContainer.innerHTML = html;
                    });

                    //this.renderPredictedIssues();
                }
            } else {
                this.renderAllIssues();

                //this.renderStartedIssues();

                this._getCompletedIssuesListData(this.context.pageContext.web.absoluteUrl, '', 'Issues').then((response) => {
                    let html : string = '';
                    for (let i : number = 0; i < response.value.length; i++) {
                        html += `<li class="ms-font-m ms-ListItem">${response.value[i].Title}</li>`;
                    }
                    const completedIssuesContainer : Element = this.domElement.querySelector('#completedIssuesContainer');
                    completedIssuesContainer.innerHTML = html;
                });
            }

        }

    }

    private renderAllIssues(): void{
        if(this.properties.showForTeam){
            var peopleList : IPropertyFieldPeople[] = [];
            if (this.properties && this.properties.people) {
                peopleList = this.properties.people;
                this._getIssuesListData(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                    const displayItems : IDisplayItem[] = [];
                    for (let i : number = 0; i < response.value.length; i++) {
                        const issueItem : IIssueItem = response.value[i];
                        displayItems.push({Label: issueItem.Title, Data: issueItem.DaysAllocated, BgColor: this.bgColors[i], HoverBgColor: this.bgColors[i]});
                    }
                    Functions.renderAllIssues(displayItems);
                });
            }
        } else {
            this._getIssuesListData(this.context.pageContext.web.absoluteUrl, '', 'Issues').then((response) => {
                const displayItems : IDisplayItem[] = [];
                for (let i : number = 0; i < response.value.length; i++) {
                    const issueItem : IIssueItem = response.value[i];
                    displayItems.push({Label: issueItem.Title, Data: issueItem.DaysAllocated, BgColor: this.bgColors[i], HoverBgColor: this.bgColors[i]});
                }
                Functions.renderAllIssues(displayItems);
            });
        }
    }

    private renderStartedIssues(): void{
        if(this.properties.showForTeam){
            var peopleList : IPropertyFieldPeople[] = [];
            if (this.properties && this.properties.people) {
                    peopleList = this.properties.people;

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
            }
        } else {
            this._getStartedIssuesListData(this.context.pageContext.web.absoluteUrl, '', 'Issues').then((response) => {
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
        }

    }

     private renderPredictedIssues(): void{
         if(this.properties.showForTeam){
            var peopleList : IPropertyFieldPeople[] = [];
            if (this.properties && this.properties.people) {
                peopleList = this.properties.people;
                this._getIssuesListDataForPred(this.context.pageContext.web.absoluteUrl, peopleList[0]['fullName'], 'Issues').then((response) => {
                        let html : string = `<thead>
                                                <tr>
                                                    <th>Issue name</th>
                                                    <th>Days allocated</th>
                                                    <th>Predicted number of days</th>
                                                </tr>
                                            </thead><tbody>`;
                        var totalPredictedDays: number = 0;
                        var totalDaysAllocated: number = 0;
                        for (let i : number = 0; i < response.value.length; i++) {
                            var issueName: string = response.value[i].Title;
                            var daysAllocated: number = response.value[i].DaysAllocated;
                            totalDaysAllocated += daysAllocated;

                            var daysUsed: number = response.value[i].DaysUsed;
                            var percentComplete: number = response.value[i].Percent_x0020_Complete * 100;
                            var predictedDays: number = daysAllocated;
                            if(daysUsed > 0 && percentComplete > 0){
                                predictedDays = (100 * daysUsed)/percentComplete;
                            }
                            totalPredictedDays += predictedDays;

                            html += `<tr>
                                        <td>${issueName}</td>
                                        <td>${daysAllocated}</td>
                                        <td>${predictedDays}</td>
                                    </tr>`;
                        }

                        html += `<tr>
                                    <td>Total</td>
                                    <td>${totalDaysAllocated}</td>
                                    <td>${totalPredictedDays}</td>
                                </tr></tbody>`;
                        const completedIssuesContainer : Element = this.domElement.querySelector('#predictionContainer');
                        completedIssuesContainer.innerHTML = html;

                        if(totalPredictedDays <= totalDaysAllocated){
                            var message: string = `${peopleList[0]['fullName']} is predicted to finish the issues in ${totalPredictedDays.toFixed(2)} days of allocated ${totalDaysAllocated} days.`;
                            swal('On track',
                                message,
                                'success');
                        } else {
                            var message: string = `${peopleList[0]['fullName']} is predicted to finish the issues in ${totalPredictedDays.toFixed(2)} days of allocated ${totalDaysAllocated} days.`;
                            swal('Might need inspiration',
                                message,
                                'warning');
                        }
                    });
            }
        } else {

            this._getIssuesListDataForPred(this.context.pageContext.web.absoluteUrl, '', 'Issues').then((response) => {
                        let html : string = `<thead>
                                                <tr>
                                                    <th>Issue name</th>
                                                    <th>Days allocated</th>
                                                    <th>Predicted number of days</th>
                                                </tr>
                                            </thead><tbody>`;
                        var totalPredictedDays: number = 0;
                        var totalDaysAllocated: number = 0;
                        for (let i : number = 0; i < response.value.length; i++) {
                            var issueName: string = response.value[i].Title;
                            var daysAllocated: number = response.value[i].DaysAllocated;
                            totalDaysAllocated += daysAllocated;

                            var daysUsed: number = response.value[i].DaysUsed;
                            var percentComplete: number = response.value[i].Percent_x0020_Complete * 100;
                            var predictedDays: number = daysAllocated;
                            if(daysUsed > 0 && percentComplete > 0){
                                predictedDays = (100 * daysUsed)/percentComplete;
                            }
                            totalPredictedDays += predictedDays;

                            html += `<tr>
                                        <td>${issueName}</td>
                                        <td>${daysAllocated}</td>
                                        <td>${predictedDays}</td>
                                    </tr>`;
                        }

                        html += `<tr>
                                    <td>Total</td>
                                    <td>${totalDaysAllocated}</td>
                                    <td>${totalPredictedDays}</td>
                                </tr></tbody>`;
                        const completedIssuesContainer : Element = this.domElement.querySelector('#predictionContainer');
                        completedIssuesContainer.innerHTML = html;

                        if(totalPredictedDays <= totalDaysAllocated){
                            var message: string = `${this.properties.teamName} team is predicted to finish the issues in ${totalPredictedDays.toFixed(2)} days of allocated ${totalDaysAllocated} days.`;
                            swal('On track',
                                message,
                                'success');
                        } else {
                            var message: string = `${this.properties.teamName} team is predicted to finish the issues in ${totalPredictedDays.toFixed(2)} days of allocated ${totalDaysAllocated} days.`;
                            swal('Might need inspiration',
                                message,
                                'warning');
                        }
                    });

        }
    }

    public render() : void {
        var peopleList: IPropertyFieldPeople[] = [];

        var teamOrPerson: string = '';
        if(this.properties.showForTeam){
            var selectedPerson: string = '';
            var selectedPersonImg: string = '';
            var jobTitle: string = '';
            if (this.properties && this.properties.people) {
                peopleList = this.properties.people;
                selectedPerson = peopleList[0]['fullName'];
                selectedPersonImg = peopleList[0]['imageUrl'];
                jobTitle = peopleList[0]['jobTitle'];
            }
            teamOrPerson = `<div class="ms-Persona">
                <div class="ms-Persona-imageArea">
                    <img class="ms-Persona-image" src="${selectedPersonImg}">
                </div>
                <div class="ms-Persona-details ${styles.minLeftPad}">
                    <div class="ms-Persona-primaryText">${selectedPerson}</div>
                    <div class="ms-Persona-secondaryText">${jobTitle}</div>
                </div>
            </div>`
        } else {
            teamOrPerson = `<span class="ms-font-l">${this.properties.teamName} team</span>`;
        }

        this.domElement.innerHTML = `
        <div id="tabs" class="${styles.issuesView}">
            <p class="ms-font-xxl">Issue status for</p>
            ${teamOrPerson}
            <br/>
            <ul>
                <li><a href="#tabs-1" class="ms-font-l" id="aAllIssues"><i class="ms-Icon ms-Icon--Error" aria-hidden="true"></i> All Issues</a></li>
                <li><a href="#tabs-2" class="ms-font-l" id="aIssuesStarted"><i class="ms-Icon ms-Icon--WorkFlow" aria-hidden="true"></i> Started</a></li>
                <li><a href="#tabs-3" class="ms-font-l" id="aIssuesCompleted"><i class="ms-Icon ms-Icon--CompletedSolid" aria-hidden="true"></i> Completed</a></li>
                <li><a href="#tabs-4" class="ms-font-l" id="aPrediction"><i class="ms-Icon ms-Icon--Health" aria-hidden="true"></i> Prediction</a></li>
                <li><a href="#tabs-5" class="ms-font-l" id="aSomedata"><i class="ms-Icon ms-Icon--ViewAll" aria-hidden="true"></i> Issues List</a></li>
            </ul>
                <div id="tabs-1">
                    <p class="ms-font-l ${styles.header}">All issues and days allocated</p>
                    <div id="allIssuesChart">
                        <canvas></canvas>
                    </div>
                </div>
                <div id="tabs-2">
                    <p class="ms-font-l ${styles.header}">Issues started</p>
                    <div id="startedIssuesChart">
                        <canvas></canvas>
                    </div>
                </div>
                <div id="tabs-3">
                    <p class="ms-font-l ${styles.header}">Issues Completed</p>
                    <ul id="completedIssuesContainer" class="ms-List"></ul>
                </div>
                    <div id="tabs-4">
                    <p class="ms-font-l ${styles.header}">Prediction</p>
                    <table class="ms-Table" id="predictionContainer">
                    </table>
                </div>
                <div id="tabs-5">
                    <p class="ms-font-l ${styles.header}">Link to all issues</p>
                    <a class="ms-Link ms-font-m" href="${this.context.pageContext.web.absoluteUrl}/Lists/Issues" target="_blank">Click here to view all issues</a>
                </div>
        </div>
`;

        this._renderIssuesAsync();
        $("#tabs").tabs();
        this.registerEvents();
    }

    private registerEvents(): void{
        const webPart: IssuesViewWebPart = this;
        $('#aAllIssues').click(function(){
            $('#allIssuesChart').html();
            $('#allIssuesChart').html('<canvas></canvas>');
            webPart.renderAllIssues();
        });
        $('#aIssuesStarted').click(function(){
            $('#startedIssuesChart').html();
            $('#startedIssuesChart').html('<canvas></canvas>');
            webPart.renderStartedIssues();
        });
        $('#aPrediction').click(function(){
            $('#predictionContainer').html();
            webPart.renderPredictedIssues();
        });

    }

    private onLoaded() : void {}

    private rendered(carousel : any) {}

    protected get propertyPaneSettings() : IPropertyPaneSettings {
        let templateProperty: any;
        if (!this.properties.showForTeam) {
            templateProperty = PropertyPaneTextField('teamName', {
                label: 'Project Team name'
            });
        } else {
            templateProperty = PropertyFieldPeoplePicker('people', {
                                    label: "Select a user",
                                    initialData: this.properties.people,
                                    allowDuplicate: false,
                                    onPropertyChange: this.onPropertyChange,
                                    context: this.context
                                })
        }
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
                                PropertyPaneToggle('showForTeam', {
                                    label: 'Show issues for team'
                                }),
                                templateProperty
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
