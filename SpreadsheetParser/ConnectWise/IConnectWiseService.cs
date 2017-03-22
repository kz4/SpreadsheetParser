using SpreadsheetParser.ConnectWise;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SpreadsheetParser
{
    public interface IConnectWiseService
    {
        Task<Ticket> AddTicket(Ticket ticket);
        Task<Ticket> CancelTicket(int ticketId);
        Task<Ticket> CloseTicket(int ticketId);
        Task<Ticket> ChangeCompany(int ticketId, string companyId);
        Task<Ticket> ChangeGenerically(int ticketId, string companyId, string operation, string path);
    }
}