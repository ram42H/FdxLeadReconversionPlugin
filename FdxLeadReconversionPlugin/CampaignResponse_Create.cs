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
                        campaignResponse.Attributes["telephone"] = Regex.Replace(campaignResponse.Attributes["telephone"].ToString(), @"[^0-9]+", "");

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
                        phone = Regex.Replace(campaignResponse.Attributes["telephone"].ToString(), @"[^0-9]+", "");
                    //phone = campaignResponse.Attributes["telephone"].ToString();
                    if (campaignResponse.Attributes.Contains("emailaddress"))
                        email = campaignResponse.Attributes["emailaddress"].ToString();

                    firstName = firstName.Replace("'", " ");
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
                    Guid activeContact = new Guid();
                    Guid inactiveContact = new Guid();
                    for (int i = 0; i < contacts.Entities.Count; i++)
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
                    for (int i = 0; i < leads.Entities.Count; i++)
                    {
                        if ((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 0) && (openLead == Guid.Empty))
                        {
                            openLead = leads.Entities[i].Id;
                        }
                        else if ((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 1) && (qualifiedLead == Guid.Empty))
                        {
                            qualifiedLead = leads.Entities[i].Id;
                        }
                        else if ((((OptionSetValue)leads.Entities[i].Attributes["statecode"]).Value == 2) && (disqualifiedLead == Guid.Empty))
                        {
                            disqualifiedLead = leads.Entities[i].Id;
                        }

                        if ((openLead != Guid.Empty) && (qualifiedLead != Guid.Empty) && (disqualifiedLead != Guid.Empty))
                            break;
                    }
                    #endregion

                    if (activeContact != Guid.Empty)
                    {
                        #region Active contact....
                        //Update flag and last touch point on the contact....
                        step = 7;
                        this.updateEntity(new Entity("contact") { Id = activeContact }, campaignResponse, service);
                        //Tag campaign response to contact....
                        step = 8;
                        //this.tagCampaignResponse(new Entity("fdx_contactcampaignresponse"), campaignResponse, activeContact, "fdx_contact", service);
                        campaignResponse["fdx_reconversioncontact"] = new EntityReference("contact", activeContact);

                        //S985 - Update Contact CR's to its Account
                        //QuerybyAttribute to get Account from contact

                      //  QueryExpression ContactQuery = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid"), new CRMQueryExpression[] { new CRMQueryExpression("contactid", ConditionOperator.Equal, activeContact)});

                        QueryByAttribute AccountQueryBycontactId = new QueryByAttribute("contact");
                        AccountQueryBycontactId.AddAttributeValue("contactid", activeContact);
                        AccountQueryBycontactId.ColumnSet = new ColumnSet("contactid", "parentcustomerid");
                        EntityCollection contactrecords = service.RetrieveMultiple(AccountQueryBycontactId);
                        EntityReference contacc = new EntityReference("account");
                        int oppCount = 0;

                        foreach (Entity contactrec in contactrecords.Entities)
                        {
                            if (contactrec.Attributes.Contains("parentcustomerid"))
                            {

                                contacc.Id = ((EntityReference)contactrec["parentcustomerid"]).Id;
                                //Update Customer on Campaign response
                                EntityReference ContAccount = new EntityReference("account", contacc.Id);

                                Entity customer = new Entity("activityparty");
                                customer.Attributes["partyid"] = ContAccount;
                                EntityCollection Customerentity = new EntityCollection();
                                Customerentity.Entities.Add(customer);
                                campaignResponse["customer"] = Customerentity;
                                tracingService.Trace("Contact CR updated to Account :" + ((EntityReference)contactrec["parentcustomerid"]).Name);

                                #region Get oldest open Opportunities where parent contact is active contact...
                                step = 15;
                                string oppXmlQuery = "<fetch top='1' ><entity name='opportunity' ><attribute name='opportunityid' /><filter type='and' ><condition attribute='statecode' operator='eq' value='0' /><condition attribute='parentcontactid' operator='eq' value='{0}' /></filter><order attribute='createdon' /></entity></fetch>";
                                query = new FetchExpression(string.Format(oppXmlQuery, activeContact.ToString()));
                                opportunities = service.RetrieveMultiple(query);
                                oppCount = opportunities.Entities.Count;
                                if (oppCount > 0)
                                {
                                    //Update flag and last touch point on the related opportunity....
                                    step = 151;
                                    this.updateEntity(new Entity("opportunity") { Id = opportunities.Entities[0].Id }, campaignResponse, service);

                                    //tag campaign response to opportunity....
                                    step = 152;
                                    campaignResponse["fdx_reconversionopportunity"] = new EntityReference("opportunity", opportunities.Entities[0].Id);
                                }
                                #endregion
                            }
                            else
                            {
                                //S985 - to handle if the opportunity's contact doesnt have an existing account and Update Opportunity CR's to its Account
                                //QuerybyAttribute to get Account from Opportunity

                                QueryByAttribute AccountQueryByoppoId = new QueryByAttribute("opportunity");
                                AccountQueryByoppoId.AddAttributeValue("opportunityid", opportunities.Entities[0].Id);
                                AccountQueryByoppoId.ColumnSet = new ColumnSet("opportunityid", "parentaccountid");
                                EntityCollection opporecords = service.RetrieveMultiple(AccountQueryByoppoId);
                                EntityReference opportacc = new EntityReference("account");

                                foreach (Entity opporec in opporecords.Entities)
                                {
                                    if (opporec.Attributes.Contains("parentaccountid"))
                                    {

                                        opportacc.Id = ((EntityReference)opporec["parentaccountid"]).Id;
                                        //Update Customer on Campaign response
                                        EntityReference OppoAccount = new EntityReference("account", opportacc.Id);

                                        Entity customer = new Entity("activityparty");
                                        customer.Attributes["partyid"] = OppoAccount;
                                        EntityCollection Customerentity = new EntityCollection();
                                        Customerentity.Entities.Add(customer);
                                        campaignResponse["customer"] = Customerentity;
                                        tracingService.Trace("Opportunity CR updated to Account :" + ((EntityReference)opporec["parentaccountid"]).Name);
                                    }
                                }

                            }
                        }
                        tracingService.Trace("Opp count :" + oppCount);
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

                            //S985 - Update Lead CR's to its Account
                            //QuerybyAttribute to get Account from Lead

                            QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
                            AccountQueryByleadId.AddAttributeValue("leadid", openLead);
                            AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
                            EntityCollection leadrecords = service.RetrieveMultiple(AccountQueryByleadId);
                            EntityReference acc = new EntityReference("account");

                            foreach (Entity leadrecord in leadrecords.Entities)
                            {
                                if (leadrecord.Attributes.Contains("parentaccountid"))
                                {
                                    acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;

                                    EntityReference LeadAccount = new EntityReference("account", acc.Id);

                                    Entity customer = new Entity("activityparty");
                                    customer.Attributes["partyid"] = LeadAccount;
                                    EntityCollection Customerentity = new EntityCollection();
                                    Customerentity.Entities.Add(customer);
                                    campaignResponse["customer"] = Customerentity;
                                    tracingService.Trace("Lead CR updated to Account :" + ((EntityReference)leadrecord["parentaccountid"]).Name);
                                }
                            }
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
                            tracingService.Trace("Opp count :" + oppCount);
                            //Create a new lead....
                            step = 96;
                            //Start - Code comment by Ram as per SMART-581...
                            //this.createdLead(campaignResponse, service);
                            //End - Code comment by Ram as per SMART-581...
                        }
                        #endregion
                    }
                    else if (inactiveContact != Guid.Empty)
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
                            //S985 - Update Lead CR's to its Account
                            //QuerybyAttribute to get Account from Lead

                            QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
                            AccountQueryByleadId.AddAttributeValue("leadid", openLead);
                            AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
                            EntityCollection leadrecords = service.RetrieveMultiple(AccountQueryByleadId);
                            EntityReference acc = new EntityReference("account");

                            foreach (Entity leadrecord in leadrecords.Entities)
                            {
                                if (leadrecord.Attributes.Contains("parentaccountid"))
                                {
                                    acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;

                                    EntityReference LeadAccount = new EntityReference("account", acc.Id);

                                    Entity customer = new Entity("activityparty");
                                    customer.Attributes["partyid"] = LeadAccount;
                                    EntityCollection Customerentity = new EntityCollection();
                                    Customerentity.Entities.Add(customer);
                                    campaignResponse["customer"] = Customerentity;
                                }
                            }

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
                    else if (openLead != Guid.Empty) //Only open lead....
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
                        //S985 - Update Lead CR's to its Account
                        //QuerybyAttribute to get Account from Lead

                        QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
                        AccountQueryByleadId.AddAttributeValue("leadid", openLead);
                        AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
                        EntityCollection leadrecords = service.RetrieveMultiple(AccountQueryByleadId);
                        EntityReference acc = new EntityReference("account");

                        foreach (Entity leadrecord in leadrecords.Entities)
                        {
                            if (leadrecord.Attributes.Contains("parentaccountid"))
                            {
                                acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;

                                EntityReference LeadAccount = new EntityReference("account", acc.Id);

                                Entity customer = new Entity("activityparty");
                                customer.Attributes["partyid"] = LeadAccount;
                                EntityCollection Customerentity = new EntityCollection();
                                Customerentity.Entities.Add(customer);
                                campaignResponse["customer"] = Customerentity;
                            }
                        }
                        #endregion
                    }
                    else if (disqualifiedLead != Guid.Empty) //Only isqualified Lead....
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
                        //S985 - Update Lead CR's to its Account
                        //QuerybyAttribute to get Account from Lead

                        QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
                        AccountQueryByleadId.AddAttributeValue("leadid", disqualifiedLead);
                        AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
                        EntityCollection leadrecords = service.RetrieveMultiple(AccountQueryByleadId);
                        EntityReference acc = new EntityReference("account");

                        foreach (Entity leadrecord in leadrecords.Entities)
                        {
                            if (leadrecord.Attributes.Contains("parentaccountid"))
                            {
                                acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;

                                EntityReference LeadAccount = new EntityReference("account", acc.Id);

                                Entity customer = new Entity("activityparty");
                                customer.Attributes["partyid"] = LeadAccount;
                                EntityCollection Customerentity = new EntityCollection();
                                Customerentity.Entities.Add(customer);
                                campaignResponse["customer"] = Customerentity;
                            }
                        }
                        #endregion
                    }
                    else //New record....
                    {
                        Guid leadGuid = Guid.Empty;

                        step = 13;
                        if (campaignResponse.Attributes.Contains("telephone") || campaignResponse.Attributes.Contains("emailaddress"))
                            leadGuid = this.createdLead(campaignResponse, service);

                        //Modified as part of S905....
                        if (!(leadGuid == Guid.Empty))
                        {
                            campaignResponse["fdx_reconversionlead"] = new EntityReference("lead", leadGuid);
                            campaignResponse["fdx_sourcecampaignresponse"] = true;
                        }
                    }

                    //Update campaign response....
                    step = 14;
                    //if (campaignResponse.Attributes.Contains("subject") && campaignResponse.Attributes.Contains("channeltypecode") && campaignResponse.Attributes.Contains("regardingobjectid") && campaignResponse.Attributes.Contains("firstname") && campaignResponse.Attributes.Contains("lastname") && campaignResponse.Attributes.Contains("telephone") && campaignResponse.Attributes.Contains("fdx_zippostalcode"))
                    //{
                    if (campaignResponse.Attributes.Contains("telephone") || campaignResponse.Attributes.Contains("emailaddress"))
                    {
                        campaignResponse["statecode"] = new OptionSetValue(1);
                        campaignResponse["statuscode"] = new OptionSetValue(2);
                        service.Update(campaignResponse);
                    }
                    //}

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
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private Guid createdLead(Entity _campaignResponse, IOrganizationService _service)
        {
            int step = 131;
            Guid leadGuid = Guid.Empty;

            try
            {
                Entity lead = new Entity("lead");
                if (_campaignResponse.Attributes.Contains("emailaddress"))
                    lead["emailaddress1"] = _campaignResponse.Attributes["emailaddress"];

                step = 132;
                if (_campaignResponse.Attributes.Contains("telephone"))
                    lead["telephone2"] = Regex.Replace(_campaignResponse.Attributes["telephone"].ToString(), @"[^0-9]+", "");
                //lead["telephone2"] = _campaignResponse.Attributes["telephone"].ToString();

                step = 133;
                if (_campaignResponse.Attributes.Contains("firstname"))
                    lead["firstname"] = _campaignResponse.Attributes["firstname"];

                step = 134;
                if (_campaignResponse.Attributes.Contains("lastname"))
                    lead["lastname"] = _campaignResponse.Attributes["lastname"];

                step = 134;
                if (_campaignResponse.Attributes.Contains("channeltypecode"))
                    lead["leadsourcecode"] = _campaignResponse.Attributes["channeltypecode"];

                step = 135;
                if (_campaignResponse.Attributes.Contains("fdx_zipcode"))
                {
                    //lead["fdx_zippostalcode"] = _campaignResponse.Attributes["fdx_zippostalcode"];
                    QueryExpression zipQuery = CRMQueryExpression.getQueryExpression("fdx_zipcode", new ColumnSet("fdx_zipcode"), new CRMQueryExpression[] { new CRMQueryExpression("fdx_zipcode", ConditionOperator.Equal, _campaignResponse.Attributes["fdx_zipcode"].ToString()) });
                    EntityCollection zipcodeCollection = _service.RetrieveMultiple(zipQuery);
                    if ((zipcodeCollection.Entities.Count) > 0)
                    {
                        step = 136;
                        Entity zipCode = zipcodeCollection.Entities[0];
                        lead["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", zipCode.Id);
                    }
                }

                step = 137;
                if (_campaignResponse.Attributes.Contains("regardingobjectid"))
                    lead["campaignid"] = _campaignResponse.Attributes["regardingobjectid"];

                step = 138;
                lead["relatedobjectid"] = new EntityReference("campaignresponse", _campaignResponse.Id);

                step = 139;
                if (_campaignResponse.Attributes.Contains("fdx_credential"))
                    lead["fdx_credential"] = _campaignResponse.Attributes["fdx_credential"];
                step = 140;
                if (_campaignResponse.Attributes.Contains("fdx_telephone1"))
                    lead["telephone1"] = Regex.Replace(_campaignResponse.Attributes["fdx_telephone1"].ToString(), @"[^0-9]+", "");
                //lead["telephone1"] = _campaignResponse.Attributes["fdx_telephone1"].ToString();
                step = 141;
                if (_campaignResponse.Attributes.Contains("fdx_jobtitlerole"))
                    lead["fdx_jobtitlerole"] = _campaignResponse.Attributes["fdx_jobtitlerole"];
                step = 142;
                if (_campaignResponse.Attributes.Contains("companyname"))
                    lead["companyname"] = _campaignResponse.Attributes["companyname"];
                step = 143;
                if (_campaignResponse.Attributes.Contains("fdx_websiteurl"))
                    lead["websiteurl"] = _campaignResponse.Attributes["fdx_websiteurl"];
                step = 144;
                if (_campaignResponse.Attributes.Contains("fdx_telephone3"))
                    lead["telephone3"] = Regex.Replace(_campaignResponse.Attributes["fdx_telephone3"].ToString(), @"[^0-9]+", "");
                //lead["telephone3"] = _campaignResponse.Attributes["fdx_telephone3"].ToString();
                step = 145;
                if (_campaignResponse.Attributes.Contains("fdx_address1_line1"))
                    lead["address1_line1"] = _campaignResponse.Attributes["fdx_address1_line1"];
                step = 146;
                if (_campaignResponse.Attributes.Contains("fdx_address1_line2"))
                    lead["address1_line2"] = _campaignResponse.Attributes["fdx_address1_line2"];
                step = 147;
                if (_campaignResponse.Attributes.Contains("fdx_address1_line3"))
                    lead["address1_line3"] = _campaignResponse.Attributes["fdx_address1_line3"];
                step = 148;
                if (_campaignResponse.Attributes.Contains("fdx_address1_city"))
                    lead["address1_city"] = _campaignResponse.Attributes["fdx_address1_city"];
                step = 149;
                if (_campaignResponse.Attributes.Contains("fdx_stateprovince"))
                    lead["fdx_stateprovince"] = _campaignResponse.Attributes["fdx_stateprovince"];
                step = 150;
                if (_campaignResponse.Attributes.Contains("fdx_address1_country"))
                    lead["address1_country"] = _campaignResponse.Attributes["fdx_address1_country"];
                step = 151;
                if (((OptionSetValue)_campaignResponse.Attributes["channeltypecode"]).Value == 4)
                    lead["ownerid"] = new EntityReference("systemuser", ((EntityReference)_campaignResponse.Attributes["ownerid"]).Id);
                step = 152;
                leadGuid = _service.Create(lead);
                return leadGuid;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(string.Format("An error occurred in the CampaignResponse_Create plug-in at Step {0}.", step), ex);
            }
        }
    }
}
