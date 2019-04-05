using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommunicatorAPI;
using System.Runtime.InteropServices;


namespace StatusTracker
{
    class Program
    {
        

        private static Messenger _messenger;
        private static SqlConnection _sqlConnection=null;

        static void Main(string[] args)
        {

            try
            {

                Console.WriteLine($"******* Program Started at {DateTime.Now}");
                _messenger = new Messenger();
                IMessengerContactAdvanced localUser =
                    (IMessengerContactAdvanced) _messenger.GetContact(_messenger.MySigninName, _messenger.MyServiceId);


                IMessengerContacts _contacts = (IMessengerContacts) _messenger.MyContacts;
                IMessengerContactAdvanced _contact;

                Console.WriteLine("Contact list for the local user:");
                for (int i = 0; i < _contacts.Count; i++)
                {
                    _contact = (IMessengerContactAdvanced) _contacts.Item(i);
                    Console.WriteLine("\t{0}", _contact.FriendlyName);
                    DisplayContactPresence(_contact);
                    InsertContactInfo((IMessengerContactAdvanced) _contact);
                }


                _messenger.OnContactStatusChange +=
                    new DMessengerEvents_OnContactStatusChangeEventHandler(_messenger_OnContactStatusChange);
                _messenger.OnMyStatusChange +=
                    new DMessengerEvents_OnMyStatusChangeEventHandler(_messenger_OnMyStatusChange);

                if (localUser != null)
                {
                    DisplayContactPresence(localUser);
                }
                Console.WriteLine("\nPress Enter key to exit the application.\n");
                Console.ReadLine();


                _messenger.OnContactStatusChange -=
                    new DMessengerEvents_OnContactStatusChangeEventHandler(_messenger_OnContactStatusChange);
                _messenger.OnMyStatusChange -=
                    new DMessengerEvents_OnMyStatusChangeEventHandler(_messenger_OnMyStatusChange);


                System.Runtime.InteropServices.Marshal.ReleaseComObject(_messenger);
                _messenger = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(localUser);
                localUser = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: An error occurred in status tracker.\n{ex}");
            }
            finally
            {
                Console.WriteLine("******** Program Ended at {0}", DateTime.Now);
            }
        }

        static void DisplayContactPresence(IMessengerContactAdvanced contact)
        {
            object[] presenceProps = (object[])contact.PresenceProperties;

            
            if (contact.IsSelf)
            {
                Console.WriteLine("Local User Presence Info for {0}:", contact.FriendlyName);

            }
            else
            {
                Console.WriteLine("Contact Presence Info for {0}:", contact.FriendlyName);
            }
            // Status or Machine State.
            Console.WriteLine("\tStatus: {0}", (MISTATUS)presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_MSTATE]);
            Console.WriteLine("\tStatus String: {0}", GetStatusString( (MISTATUS) presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_MSTATE]) );
            // Status string if status is set to custom.
            Console.WriteLine("\tCustom Status String: {0}", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_CUSTOM_STATUS_STRING]);
            // Presence or User state.
            Console.WriteLine("\tAvailability: {0}", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_AVAILABILITY]);
            Console.WriteLine("\tAvailability String: {0}", GetAvailabilityString((int)presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_AVAILABILITY]));
            // Presence note.
            Console.WriteLine("\tPresence Note: \n'{0}'", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_PRESENCE_NOTE]);
            Console.WriteLine("\tIs Blocked: {0}", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_IS_BLOCKED]);
            // OOF message for contact, if specified.
            Console.WriteLine("\tIs OOF: {0}", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_IS_OOF]);
            // Tooltip.
            Console.WriteLine("\tTool Tip: \n'{0}'\n", presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_TOOL_TIP]);


        }

        static void InsertContactInfo(IMessengerContactAdvanced contact)
        {
            try
            {

            
                object[] presenceProps = (object[])contact.PresenceProperties;
                var statusString = GetStatusString((MISTATUS) presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_MSTATE]);
                var availabilityString = GetAvailabilityString((int) presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_AVAILABILITY]).ToString();
                var status = (int) ((MISTATUS) presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_MSTATE]);
                var customStatus = presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_CUSTOM_STATUS_STRING];
                var customStatusString = (string)customStatus ?? string.Empty;
                var availability = (int)(presenceProps[(int)PRESENCE_PROPERTY.PRESENCE_PROP_AVAILABILITY]);
                var presenceNote = presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_PRESENCE_NOTE].ToString();
                var isBlocked = presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_IS_BLOCKED].ToString();
                var isOutOfOffice = presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_IS_OOF].ToString();
                var toolTip = presenceProps[(int) PRESENCE_PROPERTY.PRESENCE_PROP_TOOL_TIP].ToString();
                InsertRecordIntoDB(
                    contact.FriendlyName,
                    status,
                    statusString,
                    customStatusString,
                    availability,
                    availabilityString,
                    presenceNote,
                    isBlocked,
                    isOutOfOffice,
                    toolTip
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Exception occurred inserting contact {0}", ex));
            }
        }

        static string GetAvailabilityString(int availability)
        {
            switch (availability)
            {
                case 3000:
                    return "Available";
                case 4500:
                    return "Inactive";
                case 6000:
                    return "Busy";
                case 7500:
                    return "Busy-Idle";
                case 9000:
                    return "Do not disturb";
                case 12000:
                    return "Be right back";
                case 15000:
                    return "Away";
                case 18000:
                    return "Offline";
                default:
                    return "Indeterminate";
            }
        }

        static string GetStatusString(MISTATUS mStatus)
        {
            switch (mStatus)
            {
                case MISTATUS.MISTATUS_ALLOW_URGENT_INTERRUPTIONS:
                    return "Urgent interuptions only";
                case MISTATUS.MISTATUS_AWAY:
                    return "Away";
                case MISTATUS.MISTATUS_BE_RIGHT_BACK:
                    return "Be right back";
                case MISTATUS.MISTATUS_BUSY:
                    return "Busy";
                case MISTATUS.MISTATUS_DO_NOT_DISTURB:
                    return "Do no disturb";
                case MISTATUS.MISTATUS_IDLE:
                    return "Idle";
                case MISTATUS.MISTATUS_INVISIBLE:
                    return "Invisible";
                case MISTATUS.MISTATUS_IN_A_CONFERENCE:
                    return "In a conference";
                case MISTATUS.MISTATUS_IN_A_MEETING:
                    return "In a meeting";
                case MISTATUS.MISTATUS_LOCAL_CONNECTING_TO_SERVER:
                    return "Connecting to server";
                case MISTATUS.MISTATUS_LOCAL_DISCONNECTING_FROM_SERVER:
                    return "Disconnecting from server";
                case MISTATUS.MISTATUS_LOCAL_FINDING_SERVER:
                    return "Finding server";
                case MISTATUS.MISTATUS_LOCAL_SYNCHRONIZING_WITH_SERVER:
                    return "Synchronizing with server";
                case MISTATUS.MISTATUS_MAY_BE_AVAILABLE:
                    return "Inactive";
                case MISTATUS.MISTATUS_OFFLINE:
                    return "Offline";
                case MISTATUS.MISTATUS_ONLINE:
                    return "Online";
                case MISTATUS.MISTATUS_ON_THE_PHONE:
                    return "In a call";
                case MISTATUS.MISTATUS_OUT_OF_OFFICE:
                    return "Out of office";
                case MISTATUS.MISTATUS_OUT_TO_LUNCH:
                    return "Out to lunch";
                case MISTATUS.MISTATUS_UNKNOWN:
                    return "Unknown";
                default:
                    return "Indeterminate";
            }
        }

        static void _messenger_OnMyStatusChange(int hr, MISTATUS mMyStatus)
        {
            

            Console.WriteLine("\n***OnMyStatusChange***");
            DisplayContactPresence((IMessengerContactAdvanced)
                _messenger.GetContact(_messenger.MySigninName, _messenger.MyServiceId));
        }

        static void _messenger_OnContactStatusChange(object pMContact, MISTATUS mStatus)
        {
            
            Console.WriteLine("\n***Contact Updated Status Change at {0}", DateTime.Now.ToString());

            // Your code to work with presence goes here...
            Console.WriteLine("\n***OnContactStatusChange***");

            DisplayContactPresence((IMessengerContactAdvanced)pMContact);

            InsertContactInfo((IMessengerContactAdvanced) pMContact);
        }

        private static object syncRoot = new Object();

        static SqlConnection getSqlConnection()
        {
            if (_sqlConnection == null || (_sqlConnection != null && _sqlConnection.State == ConnectionState.Closed))
            {
                lock (syncRoot)
                {
                    //check
                    if (_sqlConnection != null && _sqlConnection.State == ConnectionState.Closed)
                    {
                        _sqlConnection.Dispose();
                        _sqlConnection = null;
                    }

                    if (_sqlConnection == null)
                    {
                        TEST_CONNECTION_STRING = @"Server=LocalDB;Integrated Security = SSPI; Initial Catalog = StatusTracker"
                        _sqlConnection = new SqlConnection(TEST_CONNECTION_STRING);
                        _sqlConnection.Open();
                    }                     

                }
            }
            
            

            return _sqlConnection;
        }
        static int InsertRecordIntoDB(
            string friendlyName,
            int status,
            string statusString,
            string customStatus,
            int availability,
            string availabilityString,
            string presenceNote,
            string isBlocked,
            string isOutOfOffice,
            string toolTip
            )
        {
            var conn = getSqlConnection();

            SqlCommand cmd = new SqlCommand(string.Format(@"INSERT INTO StatusTracker.dbo.StatusHistory 
                                            (

                                                [FriendlyName]
                                                  ,[Status]
                                                  ,[StatusString]
                                                  ,[CustomStatusString]
                                                  ,[Availability]
                                                  ,[AvailabilityString]
                                                  ,[PresenceNote]
                                                  ,[IsBlocked]
                                                  ,[IsOutOfOffice]
                                                  ,[ToolTip]
                                            )
                                            SELECT
                                            [FriendlyName] = '{0}'
                                            ,[Status]={1}
                                            ,[StatusString]='{2}'
                                            ,[CustomStatusString]='{3}'
                                            ,[Availability]={4}
                                            ,[AvailabilityString]='{5}'
                                            ,[PresenceNote]='{6}'
                                            ,[IsBlocked]='{7}'
                                            ,[IsOutOfOffice]='{8}'
                                            ,[ToolTip]='{9}'

                                             ",friendlyName,
                                             status,
                                             statusString,
                                             customStatus,
                                             availability,
                                             availabilityString,
                                             presenceNote,
                                             isBlocked,
                                             isOutOfOffice,
                                             toolTip), conn);
            return cmd.ExecuteNonQuery();
        }

    }
}
