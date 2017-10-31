import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from "@microsoft/sp-loader";

import "jquery";
import * as angular from "angular";
import "kendo";

// angular.module("ExternalWebpart", []);
angular.module("ExternalWebpart", ["kendo.directives"]);

export interface IExternalWebpartWebPartProps {
}

export default class ExternalWebpartWebPart extends BaseClientSideWebPart<IExternalWebpartWebPartProps> {

  private $injector: angular.auto.IInjectorService;

  protected onInit(): Promise<void> {
    const promise: Promise<void> = new Promise<void>((resolve: () => void, reject: () => void) => {
      resolve();
    });

    SPComponentLoader.loadCss(`https://kendo.cdn.telerik.com/2017.2.621/styles/kendo.common-office365.min.css`); 
    SPComponentLoader.loadCss(`https://kendo.cdn.telerik.com/2017.2.621/styles/kendo.office365.min.css`); 
    SPComponentLoader.loadCss(`https://kendo.cdn.telerik.com/2017.2.621/styles/kendo.office365.mobile.min.css`);

    return promise;
  }

  public render(): void {
    if (this.renderedOnce === false) {
      this.domElement.innerHTML = `
            <div>
              <p>
                Date: <input kendo-date-picker />
              </p>
              <p>
                Test angular: <input type="text" ng-model="name">
                {{ name }}
              </p>
            </div>`;

      this.$injector = angular.bootstrap(this.domElement, ["ExternalWebpart"]);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [ ]
    };
  }
}
