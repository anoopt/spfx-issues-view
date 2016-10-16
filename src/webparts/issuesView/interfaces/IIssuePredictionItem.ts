export interface IIssuePredictionItems {
    value : IIssuePredictionItem[];
}

export interface IIssuePredictionItem {
    Title : string;
    Assigned_x0020_To : any;
    Issue_x0020_Status : string;
    Percent_x0020_Complete : number;
    DaysAllocated : number;
    DaysUsed: number;
}