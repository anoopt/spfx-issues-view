import { IDisplayItem } from './interfaces/IDisplayItem'
import { IDisplayItems } from './interfaces/IDisplayItem'

export default class Functions {
  public static renderAllIssues(displayItems: IDisplayItem[] ): void{

      var canvas = <HTMLCanvasElement> $("#allIssuesChart").find('canvas').get(0);
      var ctx = <CanvasRenderingContext2D> canvas.getContext("2d");

      var labelsToBeShown: string[] = [];
      var dataToBeShown: number[] = [];
      var bgColorsToBeShown: string[] = [];
      var hoverBgColorsToBeShown: string[] = [];

      for (let i : number = 0; i < displayItems.length; i++) {
          const displayItem : IDisplayItem = displayItems[i];

          labelsToBeShown.push(displayItem.Label);
          dataToBeShown.push(displayItem.Data);
          bgColorsToBeShown.push(displayItem.BgColor);
          hoverBgColorsToBeShown.push(displayItem.HoverBgColor);
      }

      var data = {
          labels: labelsToBeShown,
          datasets: [
              {
                  label: "Days allocated",
                  data: dataToBeShown,
                  backgroundColor: bgColorsToBeShown,
                  hoverBackgroundColor: hoverBgColorsToBeShown
              }]
      };

      var options = {
          segmentShowStroke: false,
          animateRotate: true,
          animateScale: false,
          percentageInnerCutout: 50,
          tooltipTemplate: "<%= value %> days"
      }


      var allIssuesChart = new Chart(ctx, {
          type: 'doughnut',
          data: data,
          options: options
      });
      allIssuesChart.generateLegend();
  }

  public static renderIssuesStarted(displayItems: IDisplayItem[] ): void{

      var canvas = <HTMLCanvasElement> $("#startedIssuesChart").find('canvas').get(0);
      var ctx = <CanvasRenderingContext2D> canvas.getContext("2d");

      var labelsToBeShown: string[] = [];
      var dataToBeShown: number[] = [];
      var bgColorsToBeShown: string[] = [];
      var borderColorsToBeShown: string[] = [];

      for (let i : number = 0; i < displayItems.length; i++) {
          const displayItem : IDisplayItem = displayItems[i];

          labelsToBeShown.push(displayItem.Label);
          dataToBeShown.push(displayItem.Data);
          bgColorsToBeShown.push(displayItem.BgColor);
          borderColorsToBeShown.push(displayItem.HoverBgColor);
      }

      var data = {
          labels: labelsToBeShown,
          datasets: [
              {
                  label: "Percent complete",
                  data: dataToBeShown,
                  backgroundColor: bgColorsToBeShown,
                  borderColor: borderColorsToBeShown,
                  borderWidth: 1
              }]
      };

      var options = {
              scales: {
                  xAxes: [{
                      stacked: true
                  }],
                  yAxes: [{
                      stacked: true
                  }]
              }
          }

      var startedIssuesChart = new Chart(ctx, {
          type: 'bar',
          data: data,
          options: options
      });
  }
}