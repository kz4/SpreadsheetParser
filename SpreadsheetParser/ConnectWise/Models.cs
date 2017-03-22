using System;
using System.Collections.Generic;
using SpreadsheetParser.ConnectWise;

namespace SpreadsheetParser.ConnectWise
{
    public class Info
    {
        public DateTime? lastUpdated { get; set; }
        public string updatedBy { get; set; }
        public string activities_href { get; set; }
        public string timeentries_href { get; set; }
        public string scheduleentries_href { get; set; }
        public string documents_href { get; set; }
        public string products_href { get; set; }
        public string configurations_href { get; set; }
        public string tasks_href { get; set; }
        public string notes_href { get; set; }
        public string agreement_href { get; set; }
        public string source_href { get; set; }
        public string location_href { get; set; }
        public string priority_href { get; set; }
        public string image_href { get; set; }
        public string team_href { get; set; }
        public string type_href { get; set; }
        public string site_href { get; set; }
        public string company_href { get; set; }
        public string status_href { get; set; }
        public string board_href { get; set; }
    }

    public class Board
    {
        public int? id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Status
    {
        public int? id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
        public override string ToString()
        {
            return $"id: {id ?? 0}, name: {name}";
        }
    }

    public class Project
    {
        public int? id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Phase
    {
        public int? id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Company
    {
        public int? id { get; set; }
        public string identifier { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Site
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Country
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Contact
    {
        public int? id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Type
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class SubType
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Item
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Team
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Owner
    {
        public int id { get; set; }
        public string identifier { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Priority
    {
        public int id { get; set; }
        public string name { get; set; }
        public int? sort { get; set; }
        public Info _info { get; set; }
    }

    public class ServiceLocation
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Source
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Opportunity
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class Agreement
    {
        public int id { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class CustomField
    {
        public int id { get; set; }
        public string caption { get; set; }
        public string type { get; set; }
        public string entryMethod { get; set; }
        public int? numberOfDecimals { get; set; }
        public string value { get; set; }
    }

    public class Ticket
    {
        public int? id { get; set; }
        public string summary { get; set; }
        public string recordType { get; set; }
        public Board board { get; set; }
        public Status status { get; set; }
        public Project project { get; set; }
        public Phase phase { get; set; }
        public string wbsCode { get; set; }
        public Company company { get; set; }
        public Site site { get; set; }
        public string siteName { get; set; }
        public string addressLine1 { get; set; }
        public string addressLine2 { get; set; }
        public string city { get; set; }
        public string stateIdentifier { get; set; }
        public string zip { get; set; }
        public Country country { get; set; }
        public Contact contact { get; set; }
        public string contactPhoneNumber { get; set; }
        public string contactPhoneExtension { get; set; }
        public string contactEmailAddress { get; set; }
        public Type type { get; set; }
        public SubType subType { get; set; }
        public Item item { get; set; }
        public Team team { get; set; }
        public Owner owner { get; set; }
        public Priority priority { get; set; }
        public ServiceLocation serviceLocation { get; set; }
        public Source source { get; set; }
        public DateTime? requiredDate { get; set; }
        public float? budgetHours { get; set; }
        public Opportunity opportunity { get; set; }
        public Agreement agreement { get; set; }
        public string severity { get; set; }
        public string impact { get; set; }
        public string externalXRef { get; set; }
        public string poNumber { get; set; }
        public int? knowledgeBaseCategoryId { get; set; }
        public int? knowledgeBaseSubCategoryId { get; set; }
        public bool? allowAllClientsPortalView { get; set; }
        public bool? customerUpdatedFlag { get; set; }
        public bool? automaticEmailContactFlag { get; set; }
        public bool? automaticEmailResourceFlag { get; set; }
        public bool? automaticEmailCcFlag { get; set; }
        public string automaticEmailCc { get; set; }
        public string initialDescription { get; set; }
        public string initialInternalAnalysis { get; set; }
        public string initialResolution { get; set; }
        public string contactEmailLookup { get; set; }
        public bool? processNotifications { get; set; }
        public bool? skipCallback { get; set; }
        public string closedDate { get; set; }
        public string closedBy { get; set; }
        public bool? closedFlag { get; set; }
        public string dateEntered { get; set; }
        public string enteredBy { get; set; }
        public double? actualHours { get; set; }
        public bool? approved { get; set; }
        public string subBillingMethod { get; set; }
        public int? subBillingAmount { get; set; }
        public string subDateAccepted { get; set; }
        public string dateResolved { get; set; }
        public string dateResplan { get; set; }
        public string dateResponded { get; set; }
        public int? resolveMinutes { get; set; }
        public int? resPlanMinutes { get; set; }
        public int? respondMinutes { get; set; }
        public bool? isInSla { get; set; }
        public int? knowledgeBaseLinkId { get; set; }
        public string resources { get; set; }
        public int? parentTicketId { get; set; }
        public bool? hasChildTicket { get; set; }
        public string knowledgeBaseLinkType { get; set; }
        public string billTime { get; set; }
        public string billExpenses { get; set; }
        public string billProducts { get; set; }
        public string predecessorType { get; set; }
        public int? predecessorId { get; set; }
        public bool? predecessorClosedFlag { get; set; }
        public int? lagDays { get; set; }
        public bool? lagNonworkingDaysFlag { get; set; }
        public string estimatedStartDate { get; set; }
        public int? duration { get; set; }
        public Info _info { get; set; }
        public List<CustomField> customFields { get; set; }

        public override string ToString()
        {
            return $"id: {id ?? 0}, summary: {summary}, recordType: {recordType}, dateEntered: {dateEntered}";
        }
    }

    public class ResponseMessage
    {
        public string code { get; set; }
        public string message { get; set; }
        public Error[] errors { get; set; }
    }

    public class Error
    {
        public string code { get; set; }
        public string message { get; set; }
        public string resource { get; set; }
        public string field { get; set; }
    }

    public class Member
    {
        public int id { get; set; }
        public string identifier { get; set; }
        public string name { get; set; }
        public Info _info { get; set; }
    }

    public class ServiceNote
    {
        public int id { get; set; }
        public int ticketId { get; set; }
        public string text { get; set; }
        public Boolean detailDescriptionFlag { get; set; }
        public Boolean internalAnalysisFlag { get; set; }
        public Boolean resolutionFlag { get; set; }
        public Member member { get; set; }
        public Contact contact { get; set; }
        public Boolean? customerUpdatedFlag { get; set; }
        public Boolean? processNotifications { get; set; }
        public DateTime dateCreated { get; set; }
        public string createdBy { get; set; }
        public Boolean internalFlag { get; set; }
        public Boolean externalFlag { get; set; }
        public Info _info { get; set; }
    }


    public class PatchOperation
    {
        public string op { get; set; }
        public string path { get; set; }
        public object value { get; set; }

        private static PatchOperation ChangeTicketStatus(string statusValue)
        {
            return new PatchOperation
            {
                op = "replace",
                path = "status",
                value = new { Name = statusValue }
            };
        }

        public static PatchOperation CancelTicket()
        {
            return ChangeTicketStatus("Cancelled");
        }
        public static PatchOperation CloseTicket()
        {
            return ChangeTicketStatus("Closed");
        }
        public static PatchOperation NewTicket()
        {
            return ChangeTicketStatus("New");
        }
        public static PatchOperation ResolveTicket()
        {
            return ChangeTicketStatus("Resolved");
        }

        private static PatchOperation ChangeTicketCompany(string companyId)
        {
            return new PatchOperation
            {
                op = "replace",
                path = "company",
                value = new { Id = companyId }
            };
        }

        public static PatchOperation ChangeTicket(string companyId)
        {
            return ChangeTicketCompany(companyId);
        }

        private static PatchOperation ChangeGenericHelper(string companyId, string operation, string table)
        {
            return new PatchOperation
            {
                op = operation,
                path = table,
                value = new { Id = companyId }
            };
        }

        public static PatchOperation ChangeGenericOpPath(string companyId, string operation, string path)
        {
            return ChangeGenericHelper(companyId, operation, path);
        }
    }
}
