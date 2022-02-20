using System;
using System.Collections.Generic;
using System.Web;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

/// <summary>
/// ClientLessonData 的摘要描述
/// </summary>
public class ClientLessonData {
    private string _material;
    private string _sessionSN;
    public string MaterialSN { get { return _material; } }
    public string SessionSN { get { return _sessionSN; } }

    public ClientLessonData(string materialSN, string sessionSN) {
        _material = materialSN;
        _sessionSN = sessionSN;
    }
}

public class LessonDataHelper {
    public static List<ClientLessonData> GetAttendSessions(string clientSn) {
        List<ClientLessonData> sessions = HttpContext.Current.Session["ClientLessonData"] as List<ClientLessonData>;

        if (sessions == null) {
            sessions = new List<ClientLessonData>();

            using (SqlConnection cn =
                new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ConnectionString)) {
                using (SqlCommand cmd = new SqlCommand(
                    "SELECT distinct isnull(session_sn,0) as session_sn,isnull(attend_mtl_1,'') as attend_mtl_1  " +
                    "FROM dbo.client_attend_list " +
                    "WHERE (client_sn = @client_sn) AND (valid = 1) ", cn)) {
                    cn.Open();
                    cmd.Parameters.Add("client_sn", SqlDbType.Int).Value = clientSn;

                    using (SqlDataReader dr = cmd.ExecuteReader()) {
                        while (dr.Read()) {
                            sessions.Add(new ClientLessonData(
                                Convert.ToString(dr["attend_mtl_1"]),
                                Convert.ToString(dr["session_sn"])
                            ));

                        }
                    }
                }
            }
        }

        return sessions;
    }

    public static List<string> GetRecordFile(string clientSn) {
        List<string> files = HttpContext.Current.Session["ClientRecordFile"] as List<string>;

        if (files == null) {
            files = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ConnectionString)) {
                using (SqlCommand cmd = new SqlCommand(
                    "SELECT course FROM sessionrecord_view_list WHERE (client_sn =@client_sn) AND valid= 1", cn)) {
                    cn.Open();
                    cmd.Parameters.Add("client_sn", SqlDbType.Int).Value = clientSn;

                    using (SqlDataReader dr = cmd.ExecuteReader()) {
                        while (dr.Read()) {
                            files.Add(Convert.ToString(dr["course"]));
                        }
                    }
                }
            }
        }

        return files;
    }
}