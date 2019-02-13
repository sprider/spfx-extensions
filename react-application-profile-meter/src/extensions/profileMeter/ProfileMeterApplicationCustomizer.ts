import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import * as strings from 'ProfileMeterApplicationCustomizerStrings';
import * as React from "react";
import * as ReactDom from "react-dom";
import ProfileMeter from "./components/ProfileMeter";
import IProfileMeterProps from "./components/IProfileMeterProps";

const LOG_SOURCE: string = 'ProfileMeterApplicationCustomizer';

export interface IProfileMeterApplicationCustomizerProperties {
  testMessage: string;
}

export default class ProfileMeterApplicationCustomizer
  extends BaseApplicationCustomizer<IProfileMeterApplicationCustomizerProperties> {

  private _headerPlaceholder: PlaceholderContent;

  private _isHeaderReady(): boolean {
    return !this._headerPlaceholder
      && this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Top) !== -1;
  }

  private _onDispose(): void {
    console.log(`${LOG_SOURCE} Dispossed`);
  }

  private _renderPlaceHolders(): void {

    if (this._isHeaderReady()) {

      this._headerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: this._onDispose
      });

      if (!this._headerPlaceholder) {
        console.error(`${LOG_SOURCE} The expected placeholder (PageHeader) was not found.`);
        return;
      }

      if (this._headerPlaceholder.domElement) {
        const element: React.ReactElement<IProfileMeterProps> = React.createElement(
          ProfileMeter,
          {
            context: this.context
          }
        );
        ReactDom.render(element, this._headerPlaceholder.domElement);
      }

    }
  }

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();
    return Promise.resolve();
  }

}
