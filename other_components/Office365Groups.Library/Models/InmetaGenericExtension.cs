using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http;

namespace Office365Groups.Library.Models
{
    [JsonObject(MemberSerialization = MemberSerialization.OptIn)]
    public class InmetaGenericExtension
    {

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString00", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString00", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString00", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString01", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString01 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString01", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString01 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString01", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString01 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString02", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString02 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString02", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString02 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString02", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString02 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString03", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString03 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString03", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString03 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString03", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString03 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString04", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString04 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString04", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString04 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString04", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString04 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString05", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString05 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString05", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString05 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString05", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString05 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString06", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString06 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString06", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString06 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString06", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString06 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString07", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString07 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString07", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString07 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString07", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString07 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString08", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString08 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString08", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString08 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString08", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString08 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyString09", Required = Newtonsoft.Json.Required.Default)]
        public string KeyString09 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelString09", Required = Newtonsoft.Json.Required.Default)]
        public string LabelString09 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueString09", Required = Newtonsoft.Json.Required.Default)]
        public string ValueString09 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyDateTime00", Required = Newtonsoft.Json.Required.Default)]
        public string KeyDateTime00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelDateTime00", Required = Newtonsoft.Json.Required.Default)]
        public string LabelDateTime00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueDateTime00", Required = Newtonsoft.Json.Required.Default)]
        public DateTime ValueDateTime00 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "KeyDateTime01", Required = Newtonsoft.Json.Required.Default)]
        public string KeyDateTime01 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "LabelDateTime01", Required = Newtonsoft.Json.Required.Default)]
        public string LabelDateTime01 { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore, PropertyName = "ValueDateTime01", Required = Newtonsoft.Json.Required.Default)]
        public DateTime ValueDateTime01 { get; set; }

        // Parse request and map properties
        public InmetaGenericExtension(MetadataInfo req)
        {
            this.KeyString00 = "SiteType";
            this.LabelString00 = "Områdetype";
            this.ValueString00 = req.SiteType;

            this.KeyString01 = "Title";
            this.LabelString01 = req.ShortDisplayName;
            this.ValueString01 = req.DisplayName;

            this.KeyString02 = "KTDODescription";
            this.LabelString02 = "Beskrivelse";
            this.ValueString02 = req.Description;

            this.KeyString03 = "KDTOOwner";
            this.LabelString03 = req.OwnerDisplayName;
            this.ValueString03 = req.Owners[0];

            this.KeyString04 = "KTDOParentDepartment";
            this.LabelString04 = "Overordnet avdeling";
            this.ValueString04 = req.ParentDepartment;

            this.KeyString05 = "KTDOOwnedDepartment";
            this.LabelString05 = "Eiende avdeling";
            this.ValueString05 = req.OwnedDepartment;

            this.KeyString06 = "KDTOProjectNumber";
            this.LabelString06 = "Prosjektnummer";
            this.ValueString06 = req.ProjectNumber;

            this.KeyString07 = "KDTOShortName";
            this.LabelString07 = "Kortnavn";
            this.ValueString07 = req.ShortName;

            this.KeyString08 = "KDTOProjectGoal";
            this.LabelString08 = "Prosjektmål";
            this.ValueString08 = req.ProjectGoal;

            this.KeyString09 = "KDTOProjectPurpose";
            this.LabelString09 = "Prosjektets formål";
            this.ValueString09 = req.ProjectPurpose;

            this.KeyDateTime00 = "KDTOStartDate";
            this.LabelDateTime00 = "Startdato";
            this.ValueDateTime00 = req.StartDate;

            this.KeyDateTime01 = "KDTOEndDate";
            this.LabelDateTime01 = "Sluttdato";
            this.ValueDateTime01 = req.EndDate;
        }
    }
}
