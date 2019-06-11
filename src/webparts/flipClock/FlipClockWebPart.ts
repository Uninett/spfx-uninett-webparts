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
  PropertyPaneHorizontalRule,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { escape, times } from '@microsoft/sp-lodash-subset';

import styles from './FlipClockWebPart.module.scss';
require('flipclock/dist/flipclock.css');
import * as strings from 'FlipClockWebPartStrings';
import * as FlipClock from 'flipclock';

export interface IFlipClockWebPartProps {
  clockType: string;
  clockSize: string;
  fontColor: string;
  format: string;
  showSeconds: boolean;
  weekDay: number;
  timeOfDay: number;
  untilText: string;
}

export default class FlipClockWebPart extends BaseClientSideWebPart<IFlipClockWebPartProps> {

  public render(): void {

    var extraHtml = ``;
    var clockSize = 0.8;
    var fontColor = 'black';

    if (this.properties.clockType == 'countdown') {
      extraHtml = `
      <div class="${ styles.untilStyle }">
        <p>${escape(this.properties.untilText)}</p>
      </div>
      `;
    }

    switch (this.properties.clockSize) {
      case 'small':
        clockSize = 0.5;
        break;
      case 'medium':
        clockSize = 0.8;
        break;
      case 'big':
        clockSize = 1;
        break;
    }

    fontColor = this.properties.fontColor ? 'white' : 'black';

    this.domElement.innerHTML = `
      <div class="${ styles.flipClock }">
        <div class="${ styles.container }" style="zoom: ${ clockSize }; color: ${ fontColor }">
          <div class="clock ${ styles.clockStyle }"></div>
          ` + extraHtml + `
        </div>
      </div>
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

      function getTimeUntilDate(date: Date, dayOfWeek: number, time: number) {
        // gets days until selected weekday
        var resultDate = new Date(date.getTime());
        resultDate.setDate(date.getDate() + (7 + dayOfWeek - date.getDay() - 1) % 7 +1);
        
        // gets milliseconds since midnight
        var d = new Date(),
          msSinceMidnight = d.getTime() - d.setHours(0,0,0,0);
        // converts selected time of day into milliseconds
        if (time == 24) time = 0;
        var msUntilTimeOfDay = time*60*60*1000;

        resultDate.setTime(resultDate.getTime() - msSinceMidnight + msUntilTimeOfDay);

        // if day of week is today, set days remaining to 0 instead of 7
        const _MS_PER_WEEK = 1000 * 60 * 60 * 24 * 7;
        if (Math.abs(resultDate.getTime() - date.getTime()) >= _MS_PER_WEEK) resultDate.setTime(resultDate.getTime() - _MS_PER_WEEK);
        
        return resultDate;
      }
      
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let groupTitle: string;
    let clockProps1: any;
    let clockProps2: any;
    let clockProps3: any;
    
    if (this.properties.clockType != 'countdown') {
      groupTitle = 'Time settings';

      clockProps1 = PropertyPaneChoiceGroup('format', {
        label: 'Time format',
        options: [
          { key: 'TwentyFourHourClock', text: '24-hour', checked: true},
          { key: 'TwelveHourClock', text: '12-hour', checked: false}
        ]                
      });
      clockProps2 = PropertyPaneToggle('showSeconds', {
        label: 'Show seconds',
        onText: 'Yes',
        offText: 'No'
      });
      clockProps3 = PropertyPaneLabel('emptylabel', {
        text: ""
      });      
    }

    else {
      groupTitle = 'Weekly countdown settings';

      clockProps1 = PropertyPaneDropdown('weekDay', {
        label: 'Weekday',
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
      clockProps2 = PropertyPaneDropdown('timeOfDay', {
        label: 'Time of day',
        options: [
          { key: 24, text: '00:00' },
          { key: 1, text: '01:00' },
          { key: 2, text: '02:00' },
          { key: 3, text: '03:00' },
          { key: 4, text: '04:00' },
          { key: 5, text: '05:00' },
          { key: 6, text: '06:00' },
          { key: 7, text: '07:00' },
          { key: 8, text: '08:00' },
          { key: 9, text: '09:00' },
          { key: 10, text: '10:00' },
          { key: 11, text: '11:00' },
          { key: 12, text: '12:00' },
          { key: 13, text: '13:00' },
          { key: 14, text: '14:00' },
          { key: 15, text: '15:00' },
          { key: 16, text: '16:00' },
          { key: 17, text: '17:00' },
          { key: 18, text: '18:00' },
          { key: 19, text: '19:00' },
          { key: 20, text: '20:00' },
          { key: 21, text: '21:00' },
          { key: 22, text: '22:00' },
          { key: 23, text: '23:00' }
        ],
        selectedKey: 12
      });
      clockProps3 = PropertyPaneTextField('untilText', {
        label: 'Countdown description'
      });
    }


    return {
      pages: [
        {
          header: {
            description: 'Time displays the current time, while weekly countdown lets you customize a countdown for a weekly recurring event.'
          },
          groups: [
            {
              groupName: 'General settings',
              groupFields: [
                PropertyPaneChoiceGroup('clockType', {
                  label: 'Clock type',
                  options: [
                    { key: 'time', text: 'Time', checked: true},
                    { key: 'countdown', text: 'Weekly countdown', checked: false}
                  ]                
                }),
                PropertyPaneChoiceGroup('clockSize', {
                  label: 'Size',
                  options: [
                    { key: 'small', text: 'Small', checked: false},
                    { key: 'medium', text: 'Medium', checked: true},
                    { key: 'big', text: 'Big', checked: false}
                  ]                
                }),
                PropertyPaneToggle('fontColor', {
                  label: 'Text color',
                  onText: 'White',
                  offText: 'Black',
                  disabled: this.properties.clockType == 'time'
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: groupTitle,
              groupFields: [
                clockProps1,
                clockProps2,
                clockProps3
              ]
            }
          ]
        }
      ]
    };
  }
}
