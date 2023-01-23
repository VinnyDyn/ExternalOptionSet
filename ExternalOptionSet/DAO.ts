import { Option } from "./models/Option";

export class DAO {

    private _clientUrl: string;

    constructor(clientUrl: string) {
        this._clientUrl = clientUrl;
    }

    public RetrievePicklistAttributeMetadata(table: string, column: string): Option[] {
        let options = new Array<Option>();
        let req = new XMLHttpRequest();
        req.open("GET", this._clientUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='" + table + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$filter=LogicalName eq '" + column + "'&$expand=OptionSet", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    if (result !== undefined && result.value.length > 0) {
                        var options_ = result.value[0].OptionSet.Options;
                        for (let index = 0; index < options_.length; index++) {
                            const option_ = options_[index];
                            var option = new Option();
                            option.Label = option_.Label.LocalizedLabels[0].Label;
                            option.Value = option_.Value;
                            option.State = null;
                            option.Color = option_.Color;
                            options.push(option);
                        }
                    }
                }
            }
        };
        req.send();
        return options;
    }

    public RetrieveStateAttributeMetadata(table: string): Option[] {
        let options = new Array<Option>();
        let req = new XMLHttpRequest();
        req.open("GET", this._clientUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='" + table + "')/Attributes/Microsoft.Dynamics.CRM.StateAttributeMetadata?$select=LogicalName&$expand=OptionSet", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    if (result !== undefined && result.value.length > 0) {
                        var options_ = result.value[0].OptionSet.Options;
                        for (let index = 0; index < options_.length; index++) {
                            const option_ = options_[index];
                            var option = new Option();
                            option.Label = option_.Label.LocalizedLabels[0].Label;
                            option.Value = option_.Value;
                            option.State = null;
                            option.Color = option_.Color;
                            options.push(option);
                        }
                    }
                }
            }
        };
        req.send();
        return options;
    }

    public RetrieveStatusAttributeMetadata(table: string): Option[] {
        let options = new Array<Option>();
        let req = new XMLHttpRequest();
        req.open("GET", this._clientUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='" + table + "')/Attributes/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    if (result !== undefined && result.value.length > 0) {
                        var options_ = result.value[0].OptionSet.Options;
                        for (let index = 0; index < options_.length; index++) {
                            const option_ = options_[index];
                            var option = new Option();
                            option.Label = option_.Label.LocalizedLabels[0].Label;
                            option.Value = option_.Value;
                            option.State = option_.State;
                            option.Color = option_.Color;
                            options.push(option);
                        }
                    }
                }
            }
        };
        req.send();
        return options;
    }
}