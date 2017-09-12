using Microsoft.AspNet.WebHooks;
using Microsoft.AspNet.WebHooks.Payloads;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Text;
using System.Collections;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Configuration;
using System.Xml;
using WorkItemLink = Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemLink;
using Microsoft.TeamFoundation.VersionControl.Common;
using System.Configuration;

namespace VstsReceiver.WebHooks
{
    /// <summary>
    /// This handler processes WebHooks from Visual Studio Team Services and leverages the <see cref="VstsWebHookHandlerBase"/> base handler.
    /// For details about Visual Studio Team Services WebHooks, see <c>https://www.visualstudio.com/en-us/get-started/integrate/service-hooks/webhooks-and-vso-vs</c>.
    /// </summary>
    public class VstsWebHookHandler : VstsWebHookHandlerBase
    {


        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="BuildCompletedPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, BuildCompletedPayload payload)
        {
            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="TeamRoomMessagePostedPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, TeamRoomMessagePostedPayload payload)
        {
            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="WorkItemCreatedPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, WorkItemCreatedPayload payload)
        {
            string url = System.Configuration.ConfigurationManager.AppSettings["CollectionUrl"];
            //string url = WebConfigurationManager.AppSettings["CollectionUrl"];
            Uri collectionUri = new Uri(url);
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(collectionUri);
            WorkItemStore wis = tpc.GetService<WorkItemStore>();
            var project = payload.Resource.Fields.SystemTeamProject;
            if (project != null)
            {

                Project teamProject = wis.Projects[project];
                //WorkItemType workItemType = teamProject.WorkItemTypes["评审"];
                VersionControlServer vc = tpc.GetService<VersionControlServer>();
                WorkItem wid = wis.GetWorkItem(payload.Resource.Id);
                wid.AreaPath = payload.Resource.Fields.SystemTeamProject.ToString() + "\\锁定";
                string[] groupNames = new string[] { "主管工程师", "供应商", "汇众IT部门主管", "汇众科室主管" };
                var currentProject = wis.Projects.Cast<Project>().Where(p => p.Name.Equals(teamProject.Name)).First();
                ISecurityService securityService = tpc.GetService<ISecurityService>();
                SecurityNamespace securityNamespaces = securityService.GetSecurityNamespace(SecurityConstants.RepositorySecurityNamespaceGuid);
                IGroupSecurityService2 gss = tpc.GetService<IGroupSecurityService2>();

                foreach (WorkItemLink lk in wid.WorkItemLinks)
                {
                    if (lk.LinkTypeEnd.Name == "影响")
                    {
                        WorkItem brd = wis.GetWorkItem(lk.TargetId);
                        ////brd.Links.
                        brd.AreaPath = payload.Resource.Fields.SystemTeamProject.ToString() + "\\锁定";
                        brd.Save();

                        var externalLinks = brd.Links;

                        foreach (Link link in externalLinks)
                        {
                            ExternalLink exlink = link as ExternalLink;
                            if (exlink != null)
                            {
                                string artifact = System.Web.HttpUtility.UrlDecode(exlink.LinkedArtifactUri);
                                string artifact1 = System.Web.HttpUtility.UrlDecode(artifact);
                                string artifact2 = System.Web.HttpUtility.UrlDecode(artifact1);
                                //string artifact = link.LinkedArtifactUri.Ttring();
                                string path = null;
                                if (artifact2.Length > 0)
                                {
                                    path = artifact2.Substring(artifact2.IndexOf("$"), artifact2.IndexOf("&") - artifact2.IndexOf("$"));
                                }

                                string proToken = string.Format("{0}", path);
                                //IList<SecurityNamespace> securityList =(List<SecurityNamespace>)securityNamespaces.ToString();
                                IList<SecurityNamespace> securityList = new List<SecurityNamespace>();
                                securityList.Add(securityNamespaces);

                                SecurityNamespace sn = securityList.Where(s => s.Description.DisplayName == "VersionControlItems").FirstOrDefault();
                                SetSecurityToTFSGroup(gss, currentProject, sn, "主管工程师", 1, 4, proToken);
                                SetSecurityToTFSGroup(gss, currentProject, sn, "供应商", 1, 4, proToken);
                                SetSecurityToTFSGroup(gss, currentProject, sn, "汇众IT部门主管", 1, 4, proToken);
                                SetSecurityToTFSGroup(gss, currentProject, sn, "汇众科室主管", 1, 4, proToken);
                            }
                        }
                    }
                }
            }
            return Task.FromResult(true);
        }

        internal static void SetSecurityToTFSGroup(IGroupSecurityService gss, Project currentProject, SecurityNamespace sn, string groupName, int allow, int deny, string securityToken)
        {
            Identity groupIdentity = gss.ListApplicationGroups(currentProject.Uri.AbsoluteUri).Where(i => i.AccountName.Equals(groupName)).First();
            IdentityDescriptor iden = new IdentityDescriptor("Microsoft.TeamFoundation.Identity", groupIdentity.Sid);
            bool merge = true;
            sn.SetPermissions(securityToken, iden, allow, deny, merge);
        }


        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="WorkItemCommentedOnPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, WorkItemCommentedOnPayload payload)
        {
            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="CodeCheckedInPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, CodeCheckedInPayload payload)
        {
            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="WorkItemDeletedPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, WorkItemDeletedPayload payload)
        {
            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="WorkItemRestoredPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, WorkItemRestoredPayload payload)
        {

            return Task.FromResult(true);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the <see cref="WorkItemUpdatedPayload"/> WebHook.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, WorkItemUpdatedPayload payload)
        {
            string url = System.Configuration.ConfigurationManager.AppSettings["CollectionUrl"];
            Uri collectionUri = new Uri(url);
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(collectionUri);
            WorkItemStore wis = tpc.GetService<WorkItemStore>();           
            //WorkItemType workItemType = teamProject.WorkItemTypes["评审"];
            WorkItem wid = wis.GetWorkItem(payload.Resource.WorkItemId);
            if (payload.Resource.WorkItemId!=0)
            {
                wid.Open();
                bool allAgreement = true;
                //TODO:payload 里分析
                for (int i = 1; i <= 4; i++)
                {
                    string reviewer = (string)wid.Fields["Microsoft.VSTS.CMMI.RequiredAttendee" + i].Value;
                    if (reviewer != "")
                    {
                        switch ((string)wid.Fields["huizhong.VSTS.CMMI.Review" + i].Value)
                        {
                            case "通过":
                                break;
                            default:
                                allAgreement = false;
                                break;
                        }
                    }
                }
                if (allAgreement == true)
                {
                    wid.State = "通过";
                }
                else
                {
                    wid.State = "未通过";
                }
                wid.Save();
                return Task.FromResult(true);
            }
           
            return Task.FromResult(false);
        }

        /// <summary>
        /// We use <see cref="VstsWebHookHandlerBase"/> so just have to override the methods we want to process WebHooks for.
        /// This one processes the payload for unknown <c>eventType</c>.
        /// </summary>
        public override Task ExecuteAsync(WebHookHandlerContext context, JObject payload)
        {
            return Task.FromResult(true);
        }
    }
}
