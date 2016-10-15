import {IIssueItem} from './IssuesViewWebPart';

export default class MockHttpClient {

  private static _items : IIssueItem[] = [
    {
      Title: "Mock Issue 1",
      Assigned_x0020_To: {
        Title: 'Olivier Carpentier'
      },
      Issue_x0020_Status: 'In progress',
      Percent_x0020_Complete: 25,
      DaysAllocated: 2
    }, {
      Title: "Mock Issue 2",
      Assigned_x0020_To: {
        Title: 'Test'
      },
      Issue_x0020_Status: 'Complete',
      Percent_x0020_Complete: 100,
      DaysAllocated: 2
    }, {
      Title: "Mock Issue 3",
      Assigned_x0020_To: {
        Title: 'Olivier Carpentier'
      },
      Issue_x0020_Status: 'Not Started',
      Percent_x0020_Complete: 10,
      DaysAllocated: 2
    }, {
      Title: "Mock Issue 3",
      Assigned_x0020_To: {
        Title: 'Test'
      },
      Issue_x0020_Status: 'Not Started',
      Percent_x0020_Complete: 10,
      DaysAllocated: 2
    }
  ];

  public static get(restUrl : string, options?: any) : Promise < IIssueItem[] > {
    return new Promise < IIssueItem[] > ((resolve) => {
      resolve(MockHttpClient._items.filter(MockHttpClient.matchName));
    });
  }

  public static matchName(item: IIssueItem) : boolean{
      if(item.Assigned_x0020_To.Title == 'Olivier Carpentier'){
        return true;
      }
  }
}