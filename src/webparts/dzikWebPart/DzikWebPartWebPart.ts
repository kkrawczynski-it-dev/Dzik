import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DzikWebPartWebPart.module.scss';
import * as strings from 'DzikWebPartWebPartStrings';

export interface IDzikWebPartWebPartProps {
  description: string;
  dzikSlider: number;
}

export default class DzikWebPartWebPart extends BaseClientSideWebPart<IDzikWebPartWebPartProps> {

  public render(): void {
    
    function powerOfTwo(sliderInput: number): number{
      return sliderInput*sliderInput;
    }


    this.domElement.innerHTML = `
      <div class="${ styles.dzikWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to Dzik Web Part!</span>
              <p class="${ styles.subTitle }">Click pencil to edit fields</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">Value that comes from the slider: ${escape(this.properties.dzikSlider.toString())}</p>
              <p class="${ styles.description}">Value <sup>2</sup>: ${powerOfTwo(this.properties.dzikSlider)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
                PropertyPaneSlider("dzikSlider",{min:0,max:10})
              ]
            }
          ]
        }
      ]
    };
  }
}
