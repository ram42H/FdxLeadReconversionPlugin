using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FdxLeadReconversionPlugin
{
    public class CampaignResponse_Create : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            int step = 0;

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity campaignResponse = (Entity)context.InputParameters["Target"];

                if (campaignResponse.LogicalName != "campaignresponse")
                    return;

                try
                {
                    step = 1;
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    if (campaignResponse.Attributes.Contains("telephone"))
                        campaignResponse.Attributes["telephone"] = Regex.Replace(campaignResponse.Attributes["telephone"].ToString(),@"[^0-9]+","");

                    if (campaignResponse.Attributes.Contains("fdx_telephone1"))
                        campaignResponse.Attributes["fdx_telephone1"] = Regex.Replace(campaignResponse.Attributes["fdx_telephone1"].ToString(), @"[^0-9]+", "");

                    if (campaignResponse.Attributes.Contains("fdx_telephone3"))
                        campaignResponse.Attributes["fdx_telephone3"] = Regex.Replace(campaignResponse.Attributes["fdx_telephone3"].ToString(), @"[^0-9]+", "");

                    step = 2;
                    Entity contact = new Entity();
                    Entity lead = new Entity();
                    EntityCollection contacts = new EntityCollection();
                    EntityCollection leads = new EntityCollection();
                    EntityCollection opportunities = new EntityCollection();
                    FetchExpression query;
                    string firstName = "";
                    string lastName = "";
                    string phone = "";
                    string email = "";
                    if (campaignResponse.Attributes.Contains("firstname"))
                        firstName = campaignResponse.Attributes["firstname"].ToString();
                    if (campaignResponse.Attributes.Contains("lastname"))
                        lastName = campaignResponse.Attributes["lastname"].ToString();
                    if (campaignResponse.Attributes.Contains("telephone"))
                        phone = Regex.Replace(campaignResponse.Attributes["telephone"].ToString(),@"[^0-9]+", "");
                        //phone = campaignResponse.Attributes["telephone"].ToString();
                    if (campaignResponse.Attributes.Contains("emailaddress"))
                        email = campaignResponse.Attributes["emailaddress"].ToString();

                    firstName =  firstName.Replace("'", " ");
                    lastName = lastName.Replace("'", " ");

                    #region Get Contact and Lead collection....
                    //fetch xml query for lead matching....
                    string leadXmlQuery = "<fetch top='5000' ><entity name='lead' ><all-attributes/><filter type='or' ><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='telephone1' operator='eq' value='{1}' /></filter><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='telephone2' operator='eq' value='{1}' /></filter><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='emailaddress1' operator='eq' value='{2}' /></filter></filter><order attribute='createdon' /></entity></fetch>";

                    //fetch xml query for contact matching...
                    string contactXmlQuery = "<fetch top='5000' ><entity name='contact' ><all-attributes/><filter type='or' ><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='telephone1' operator='eq' value='{1}' /></filter><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='telephone2' operator='eq' value='{1}' /></filter><filter type='and' ><condition attribute='fullname' operator='eq' value='{0}' /><condition attribute='emailaddress1' operator='eq' value='{2}' /></filter></filter>    <order attribute='createdon' descending='true' /></entity></fetch>";

                    //fetch leads....
                    step = 3;
                    query = new FetchExpression(string.Format(leadXmlQuery, string.Format("{0} {1}", firstName, lastName), phone, email));
                    leads = service.RetrieveMultiple(query);

                    //fetch contacts....
                    step = 4;
                    query = new FetchExpression(string.Format(contactXmlQuery, string.Format("{0} {1}", firstName, lastName), phone, email));
                    contacts = service.RetrieveMultiple(query);
                    #endregion

                    #region Get Contacts status wise....
                    step = 5;
                    Guid activeContact = new Guid ();
                    Guid inactiveContact = new Guid();
                    for(int i = 0; i < contacts.Entities.Count; i++)
                    {
                        if ((((OptionSetValue)contacts.Entities[i].Attributes["statecode"]).Value == 0) && (activeContact == Guid.Empty))
                        {
                            activeContact = contacts.Entities[i].Id;
                        }
                        else if ((((OptionSetValue)contacts.Entities[i].Attributes["statecode"]).Value == 1) && (inactiveContact == Guid.Empty))
                        {
                            inactiveContact = contacts.Entities[i].Id;
                        }

                        if ((activeContact != Guid.Empty) && (inactiveContact != Guid.Empty))
                            break;
                    }
                    #endregion

                    #region Get Leads status wise....
                    step = 6;
                    Guid openLead = new Guid();
                    Guid qualifiedLead = new Guid();
                    Guid disqualifiedLead = new Guid();
                    for(int i = 0; i < leads.Entities.Count; i++)
                    {
                        if ((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 0) && (openLead == Guid.Empty))
                        {
                            openLead = leads.Entities[i].Id;
                        }
                        else if((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 1) && (qualifiedLead == Guid.Empty))
                        {
                            qualifiedLead = leads.Entities[i].Id;
                        }
                        else if((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 2) && (disqualifiedLead == Guid.Empty))
                        {
                            disqualifiedLead = leads.Entities[i].Id;
                        }

                        if ((openLead != Guid.Empty) && (qualifiedLead != Guid.Empty) && (disqualifiedLead != Guid.Empty))
                            break;
                    }
                    #endregion

                    if(activeContact != Guid.Empty)
                    {
                        #region Active contact....
                        //Update flag and last touch point on the contact....
                        step = 7;
                        this.updateEntity(new Entity("contact") { Id = activeContact }, campaignResponse, service);
                        //Tag campaign response to contact....
                        step = 8;
                        //this.tagCampaignResponse(new Entity("fdx_contactcampaignresponse"), campaignResponse, activeContact, "fdx_contact", service);
                        campaignResponse["fdx_reconversioncontact"] = new EntityReference("contact", activeContact);

                        #region Get oldest open Opportunities where parent contact is active contact...
                        step = 15;
                        string oppXmlQuery = "<fetch top='1' ><entity name='opportunity' ><attribute name='opportunityid' /><filter type='and' ><condition attribute='statecode' operator='eq' value='0' /><condition attribute='parentcontactid' operator='eq' value='{0}' /></filter><order attribute='createdon' /></entity></fetch>";
                        query = new FetchExpression(string.Format(oppXmlQuery, activeContact.ToString()));
                        opportunities = service.RetrieveMultiple(query);
                        int oppCount = opportunities.Entities.Count;
                        if(oppCount > 0)
                        {
                            //Update flag and last touch point on the related opportunity....
                            step = 151;
                            this.updateEntity(new Entity("opportunity") { Id = opportunities.Entities[0].Id }, campaignResponse, service);

                            //tag campaign response to opportunity....
                            step = 152;
                            campaignResponse["fdx_reconversionopportunity"] = new EntityReference("opportunity", opportunities.Entities[0].Id);
                        }
                        
                        #endregion

                        step = 9;
                        if (openLead != Guid.Empty) //If there is a open lead....
                        {
                            //Update flag and last touch point on lead....
                            step = 91;
                            this.updateEntity(new Entity("lead") { Id = openLead }, campaignResponse, service);
                            //Tag campaign response to lead....
                            step = 92;
                            //this.tagCampaignResponse(new Entity("fdx_leadcampaignresponse"), campaignResponse, openLead, "fdx_lead", service);
                            campaignResponse["fdx_reconversionlead"] = new EntityReference("lead", openLead);
                        }
                        #region (code comment) for qualified Lead scenario....
                        //else if (qualifiedLead != Guid.Empty) //If there is a qualified lead....
                        //{
                        //    //Get related opportunity....
                        //    step = 93;
                        //    QueryExpression oppQuery = CRMQueryExpression.getQueryExpression("opportunity", new ColumnSet("opportunityid"), new CRMQueryExpression[] { new CRMQueryExpression("originatingleadid", ConditionOperator.Equal, qualifiedLead) });
                        //    EntityCollection oppCollection = service.RetrieveMultiple(oppQuery);
                        //    if (oppCollection.Entities.Count > 0)
                        //    {
                        //        //Update flag and last touch point on the related opportunity....
                        //        step = 94;
                        //        this.updateEntity(new Entity("opportunity") { Id = oppCollection.Entities[0].Id }, campaignResponse, service);
                        //        //tag campaign response to opportunity....
                        //        step = 95;
                        //        //this.tagCampaignResponse(new Entity("fdx_opportunitycampaignresponse"), campaignResponse, oppCollection.Entities[0].Id, "fdx_opportunity", service);
                        //        campaignResponse["fdx_reconversionopportunity"] = new EntityReference("opportunity", oppCollection.Entities[0].Id);
                        //    }
                        //}
                        #endregion
                        else if (disqualifiedLead == Guid.Empty && oppCount == 0) //If there is no disqualifed lead....
                        {
                            //Create a new lead....
                            step = 96;
                            //Start - Code comment by Ram as per SMART-581...
                            //this.createdLead(campaignResponse, service);
                            //End - Code comment by Ram as per SMART-581...
                        }
                        #endregion
                    }
                    else if(inactiveContact != Guid.Empty)
                    {
                        #region Inactive contact....
                        step = 10;
                        if (openLead != Guid.Empty) //If there is a open lead....
                        {
                            //Update flag and last touch point on lead....
                            step = 101;
                            this.updateEntity(new Entity("lead") { Id = openLead }, campaignResponse, service);
                            //Tag campaign response to lead....
                            step = 102;
                            //this.tagCampaignResponse(new Entity("fdx_leadcampaignresponse"), campaignResponse, openLead, "fdx_lead", service);
                            campaignResponse["fdx_reconversionlead"] = new EntityReference("lead", openLead);
                        }
                        else //Create a new lead....
                        {
                            step = 103;
                            //Start - Code comment by Ram as per SMART-581...
                            //this.createdLead(campaignResponse, service);
                            //End - Code comment by Ram as per SMART-581...
                        }
                        #endregion....
                    }
                    else if(openLead != Guid.Empty) //Only open lead....
                    {
                        #region Only Open Lead....
                        step = 11;
                        //Update flag and last touch point on lead....
                        step = 111;
                        this.updateEntity(new Entity("lead") { Id = openLead }, campaignResponse, service);
                        //Tag campaign response to lead....
                        step = 112;
                        //this.tagCampaignResponse(new Entity("fdx_leadcampaignresponse"), campaignResponse, openLead, "fdx_lead", service);
                        campaignResponse["fdx_reconversionlead"] = new EntityReference("lead", openLead);
                        #endregion
                    }
                    else if(disqualifiedLead != Guid.Empty) //Only isqualified Lead....
                    {
                        #region Only disqualified Lead....
                        step = 12;
                        //Update flag and last touch point on lead....
                        step = 121;
                        this.updateEntity(new Entity("lead") { Id = disqualifiedLead }, campaignResponse, service);
                        //Tag campaign response to lead....
                        step = 122;
                        //this.tagCampaignResponse(new Entity("fdx_leadcampaignresponse"), campaignResponse, disqualifiedLead, "fdx_lead", service);
                        campaignResponse["fdx_reconversionlead"] = new EntityReference("lead", disqualifiedLead);
                        #endregion
                    }
                    else //New record....
                    {
                        step = 13;
                        this.createdLead(campaignResponse, service);
                    }

                    //Update campaign response....
                    step = 14;
                    campaignResponse["statecode"] = new OptionSetValue(1);
                    campaignResponse["statuscode"] = new OptionSetValue(2);
                    service.Update(campaignResponse);

                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the CampaignResponse_Create plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("CampaignResponse_Create: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }
        private void updateEntity(Entity _destinationEntity, Entity _campaignResponse, IOrganizationService _service)
        {
            try
            {
                _destinationEntity["fdx_reconversion"] = true;//new OptionSetValue(1);
                _destinationEntity["fdx_lasttouchpoint"] = _campaignResponse.Attributes["regardingobjectid"];
                _destinationEntity["fdx_lasttouchpointdate"] = _campaignResponse.Attributes["receivedon"];
                _service.Update(_destinationEntity);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void tagCampaignResponse(Entity _destinationEntity, Entity _campaignResponse, Guid _sourceid, string _sourceName, IOrganizationService _service)
        {
            try
            {
                _destinationEntity["fdx_campaign"] = _campaignResponse.Attributes["regardingobjectid"];
                _destinationEntity["fdx_campaignresponse"] = _campaignResponse.Id;
                _destinationEntity[_sourceName] = _sourceid;
                _service.Create(_destinationEntity);
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        private void createdLead(Entity _campaignResponse, IOrganizationService _service)
        {
            try
            {
                Entity lead = new Entity("lead");
                if (_campaignResponse.Attributes.Contains("emailaddress"))
                    lead["emailaddress1"] = _campaignResponse.Attributes["emailaddress"];
                lead["telephone2"] = Regex.Replace(_campaignResponse.Attributes["telephone"].ToString(),@"[^0-9]+", "");
                //lead["telephone2"] = _campaignResponse.Attributes["telephone"].ToString();
                lead["firstname"] = _campaignResponse.Attributes["firstname"];
                lead["lastname"] = _campaignResponse.Attributes["lastname"];
                lead["leadsourcecode"] = _campaignResponse.Attributes["channeltypecode"];
                lead["fdx_zippostalcode"] = _campaignResponse.Attributes["fdx_zippostalcode"];
                lead["campaignid"] = _campaignResponse.Attributes["regardingobjectid"];
                lead["relatedobjectid"] = new EntityReference("campaignresponse",_campaignResponse.Id);
                if (_campaignResponse.Attributes.Contains("fdx_credential"))
                    lead["fdx_credential"] = _campaignResponse.Attributes["fdx_credential"];
                if (_campaignResponse.Attributes.Contains("fdx_telephone1"))
                    lead["telephone1"] = Regex.Replace(_campaignResponse.Attributes["fdx_telephone1"].ToString(),@"[^0-9]+", "");
                    //lead["telephone1"] = _campaignResponse.Attributes["fdx_telephone1"].ToString();
                if (_campaignResponse.Attributes.Contains("fdx_jobtitlerole"))
                    lead["fdx_jobtitlerole"] = _campaignResponse.Attributes["fdx_jobtitlerole"];
                if (_campaignResponse.Attributes.Contains("companyname"))
                    lead["companyname"] = _campaignResponse.Attributes["companyname"];
                if (_campaignResponse.Attributes.Contains("fdx_websiteurl"))
                    lead["websiteurl"] = _campaignResponse.Attributes["fdx_websiteurl"];
                if (_campaignResponse.Attributes.Contains("fdx_telephone3"))
                    lead["telephone3"] = Regex.Replace(_campaignResponse.Attributes["fdx_telephone3"].ToString(),@"[^0-9]+", "");
                    //lead["telephone3"] = _campaignResponse.Attributes["fdx_telephone3"].ToString();
                if (_campaignResponse.Attributes.Contains("fdx_address1_line1"))
                    lead["address1_line1"] = _campaignResponse.Attributes["fdx_address1_line1"];
                if (_campaignResponse.Attributes.Contains("fdx_address1_line2"))
                    lead["address1_line2"] = _campaignResponse.Attributes["fdx_address1_line2"];
                if (_campaignResponse.Attributes.Contains("fdx_address1_line3"))
                    lead["address1_line3"] = _campaignResponse.Attributes["fdx_address1_line3"];
                if (_campaignResponse.Attributes.Contains("fdx_address1_city"))
                    lead["address1_city"] = _campaignResponse.Attributes["fdx_address1_city"];
                if (_campaignResponse.Attributes.Contains("fdx_stateprovince"))
                    lead["fdx_stateprovince"] = _campaignResponse.Attributes["fdx_stateprovince"];
                if (_campaignResponse.Attributes.Contains("fdx_address1_country"))
                    lead["address1_country"] = _campaignResponse.Attributes["fdx_address1_country"];

                if (((OptionSetValue)_campaignResponse.Attributes["channeltypecode"]).Value == 4)
                    lead["ownerid"] = new EntityReference("systemuser", ((EntityReference)_campaignResponse.Attributes["ownerid"]).Id);

                _service.Create(lead);
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
