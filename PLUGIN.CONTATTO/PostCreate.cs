using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace PLUGIN.CONTATTO
{
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)context.InputParameters["Target"];

                    if (true)  // if da gestire
                    {

                        Entity evento = new Entity("aicgusto_evento");


                        string fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                                            "  <entity name='aicgusto_evento'>" +
                                            "    <attribute name='aicgusto_eventoid' />" +
                                            "    <attribute name='aicgusto_name' />" +
                                            "    <attribute name='aicgusto_data' />" +
                                            "    <order attribute='aicgusto_data' descending='false' />" +
                                            "    <filter type='and'>" +
                                            "      <condition attribute='aicgusto_data' operator='this-month' />" +
                                            "      < condition attribute = 'aicgusto_data' operator= 'on-or-after' value = " + DateTime.Now + " ' />" +
                                            "    </filter>" +
                                            "  </entity>" +
                                            "</fetch>";

                        EntityCollection eventColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (eventColl != null && eventColl.Entities.Count > 0)
                        {
                            evento = eventColl.Entities.First();
                        }

                        //creazione partecipazione riferita al cliente
                        string nomeP = target.GetAttributeValue<string>("lastname");
                        string thisMonth = DateTime.Now.ToString("MMMM", new CultureInfo("it-IT"));
                        tracingService.Trace($"dopo  fetch {nomeP} \n  {evento.Id} ");
                        string nomePartecip = $"Partecipazione di {thisMonth} del signor {nomeP}";

                        //creazione partecipazione riferita al cliente
                        Entity partecipazione = new Entity("aicgusto_partecipazione");
                        partecipazione.Attributes.Add("aicgusto_name", nomePartecip);
                        partecipazione.Attributes.Add("aicgusto_evento", new EntityReference(target.LogicalName, evento.Id));
                        partecipazione.Attributes.Add("aicgusto_contatto", new EntityReference(target.LogicalName, target.Id));
                        tracingService.Trace($"dopo aggiunta attributi partecipazione {evento.Id} ");
                        Guid partecipazioneId = service.Create(partecipazione);
                        tracingService.Trace($"dopo aggiunta attributi partecipazione {evento.Id} ");


                    }

                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException($"ERRORE PostCreate : {e.Message}");
            }
        }
    }
}


