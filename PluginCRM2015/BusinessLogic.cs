using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PluginCRM2015
{
    public class BusinessLogic : IPlugin
    {
          public void Execute(IServiceProvider serviceProvider)
        {
            bool fError = false;
            try
            {
                var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity)) return;
                var entityNew = context.InputParameters["Target"] as Entity;
                if (entityNew.LogicalName != "contact") return;

                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(context.UserId);

                if (entityNew.Contains("ivg_contactnumber"))
                {
                    #region nhap tay
                    var ivg_contactnumber = entityNew.Attributes["ivg_contactnumber"].ToString();
                    if (CheckDuplicateNumber(service, ivg_contactnumber))
                    {
                        fError = true;
                        throw new InvalidPluginExecutionException("Trung contact number!");
                    }
                    #endregion
                }
                else
                {
                    #region Tu Sinh
                    GetContactNumber(service, entityNew);
                    #endregion
                }

            }
            catch (Exception ex)
            {
                if (fError)
                    throw new InvalidPluginExecutionException("Trung contact number kiem tra lai!");
                else
                    throw new Exception("Lỗi xảy ra:" + " " + ex.Message);
            }
        }

        bool CheckDuplicateNumber(IOrganizationService service, string ivg_contactnumber)
        {
            var qe = new QueryExpression("contact");
            string[] col = { "ivg_contactnumber" };
            qe.ColumnSet = new ColumnSet(col);

            ConditionExpression con = new ConditionExpression();
            con.AttributeName = "ivg_contactnumber";
            con.Operator = ConditionOperator.Equal;
            con.Values.Add(ivg_contactnumber);

            qe.Criteria.Conditions.Add(con);

            EntityCollection ens = service.RetrieveMultiple(qe);
            if (ens != null && ens.Entities.Count > 0) return true;
            return false;
        }


        private void GetContactNumber(IOrganizationService service, Entity entity)
        {
            string contractNumber = GetNumber(service);
            if (!string.IsNullOrEmpty(contractNumber))
            {
                entity.Attributes["ivg_contactnumber"] = contractNumber;
            }
            else
            {
                throw new Exception("Error Auto generate Contact Number!".ToUpper());
            }
        }

        private string GetNumber(IOrganizationService service)
        {
            int suffix = 0;
            string contractNumber = string.Empty;
            QueryExpression query = new QueryExpression("ivg_autonumber");

            FilterExpression filter = new FilterExpression(LogicalOperator.And);

            filter.AddCondition("ivg_contactnumber", ConditionOperator.Equal, "contractnumber");
            filter.AddCondition("ivg_year", ConditionOperator.Equal, DateTime.Now.Year.ToString().Substring(2));
            filter.AddCondition("ivg_month", ConditionOperator.Equal, DateTime.Now.Month.ToString());
            query.Criteria.AddFilter(filter);

            query.ColumnSet = new ColumnSet(true);
            EntityCollection collection = null;
            try
            {
                collection = service.RetrieveMultiple(query);

                if (collection != null && collection.Entities.Count > 0)
                {
                    contractNumber = null;
                    string MM = collection.Entities[0].Attributes["ivg_month"].ToString();
                    if (MM.Length == 1)
                        MM = "0" + MM;
                    suffix = int.Parse(collection.Entities[0].Attributes["ivg_number"].ToString()) + 1;
                    string suffixString = suffix.ToString("0000");
                    //if (suffixString.Length == 1)
                    //    suffixString = "000" + suffixString;
                    //if (suffixString.Length == 2)
                    //    suffixString = "00" + suffixString;
                    //if (suffixString.Length == 3)
                    //    suffixString = "0" + suffixString;
                    contractNumber = collection.Entities[0].Attributes["ivg_year"].ToString() + MM + "-" + suffixString;

                    #region update auto number entity
                    Guid autonumberId = collection.Entities[0].Id;
                    Entity ivg_autonumber = new Entity("ivg_autonumber");
                    ivg_autonumber.Id = autonumberId;
                    ivg_autonumber.Attributes["ivg_number"] = suffix.ToString();

                    service.Update(ivg_autonumber);

                    #endregion
                }
                else
                {
                    suffix = 1;
                    Entity ivg_autonumber = new Entity("ivg_autonumber");
                    ivg_autonumber.Attributes["ivg_number"] = suffix.ToString();
                    ivg_autonumber.Attributes["ivg_contactnumber"] = "contractnumber";
                    ivg_autonumber.Attributes["ivg_year"] = DateTime.Now.Year.ToString().Substring(2);
                    ivg_autonumber.Attributes["ivg_month"] = DateTime.Now.Month.ToString();

                    service.Create(ivg_autonumber);
                    string MM = DateTime.Now.Month.ToString();
                    if (MM.Length == 1)
                        MM = "0" + MM;
                    contractNumber = DateTime.Now.Year.ToString().Substring(2) + MM + "-0001";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.StackTrace);
            }
            return contractNumber;
        }

    }
}
