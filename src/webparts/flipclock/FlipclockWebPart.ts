import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel,
  IPropertyPaneGroup,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

require('flipclock/dist/flipclock.css');
import * as strings from 'FlipclockWebPartStrings';
import * as FlipClock from 'flipclock';

export interface IFlipClockWebPartProps {
  clockType: string;
  format: string;
  showSeconds: boolean;
  weekDay: number;
  timeOfDay: number;
}

export default class FlipClockWebPart extends BaseClientSideWebPart<IFlipClockWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="clock"></div>
      `;

      const el = document.querySelector('.clock');
      var faceFormat: string;      

      if (this.properties.clockType != 'countdown') {
        // sets clock format, TwentyFourHour is default value
        faceFormat = this.properties.format || 'TwentyFourHourClock';
        const clock = new FlipClock(el, {
          face: faceFormat,
          showSeconds:  this.properties.showSeconds
        });
      }
      else {
        var today = new Date;
        var weekDay = this.properties.weekDay || 5;
        var timeOfDay = this.properties.timeOfDay || 12;

        var nextFriday = getTimeUntilDate(today, weekDay, timeOfDay);
        const clock = new FlipClock(el, nextFriday, {
          face: 'DayCounter',
          countdown: true
        });
      }

      function getTimeUntilDate(date: Date, dayOfWeek: number, timeOfDay:  number) {
        // gets days until selected weekday
        var resultDate = new Date(date.getTime());0
        resultDate.setDate(date.getDate() + (7 + dayOfWeek - date.getDay() - 1) % 7 +1);
        // gets milliseconds since midnight
        var d = new Date(),
          msSinceMidnight = d.getTime() - d.setHours(0,0,0,0);
        // converts selected time of day into milliseconds
        var msUntilTimeOfDay = timeOfDay*60*60*1000;

        resultDate.setTime(resultDate.getTime() - msSinceMidnight + msUntilTimeOfDay);
        return resultDate;
      }
      
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let clockProps: any;    
    let typeName: string;

    if (this.properties.clockType != 'countdown') {
      clockProps = PropertyPaneChoiceGroup('format', {
        label: 'Time format',
        options: [
          { key: 'TwentyFourHourClock', text: '24-hour', checked: true},
          { key: 'TwelveHourClock', text: '12-hour', checked: false}
        ]                
      });
      typeName = 'Time clock properties';
    }
    else {
      clockProps = PropertyPaneDropdown('weekDay', {
        label: 'Select weekday',
        options: [
          { key: 1, text: 'Monday' },
          { key: 2, text: 'Tuesday' },
          { key: 3, text: 'Wednesday' },
          { key: 4, text: 'Thursday' },
          { key: 5, text: 'Friday' },
          { key: 6, text: 'Saturday' },
          { key: 7, text: 'Sunday' }
        ],
        selectedKey: 5
      });
      typeName = 'Countdown clock properties';

    }


    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: '',
              groupFields: [
                PropertyPaneChoiceGroup('clockType', {
                  label: 'Clock type',
                  options: [
                    { key: 'time', text: 'Time', checked: true},
                    { key: 'countdown', text: 'Countdown', checked: false}
                  ]                
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: '',
              groupFields: [
                clockProps,
                PropertyPaneToggle('showSeconds', {
                  label: 'Show seconds',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            }
             /* {
              groupName: 'Clock options',
              groupFields: [
                PropertyPaneChoiceGroup('format', {
                  label: 'Time format',
                  options: [
                    { key: 'TwentyFourHourClock', text: '24-hour', checked: true},
                    { key: 'TwelveHourClock', text: '12-hour', checked: false}
                  ]                
                }),
                PropertyPaneToggle('showSeconds', {
                  label: 'Show seconds',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: 'Countdown options',
              groupFields: [
                PropertyPaneChoiceGroup('format', {
                  label: 'Time format',
                  options: [
                    { key: 'TwentyFourHourClock', text: '24-hour', checked: true},
                    { key: 'TwelveHourClock', text: '12-hour', checked: false}
                  ]                
                }),
                PropertyPaneToggle('showSeconds', {
                  label: 'Show seconds',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            } */
          ]
        }
      ]
    };
  }
}
