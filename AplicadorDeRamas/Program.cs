using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;
using Microsoft.Office.Server.Search.Query;
using System.IO;
using Microsoft.Office.Server.Search.Administration;
using Microsoft.SharePoint.Administration;

namespace AplicadorDeRamas
{
    class Program
    {
        private static string url;
        private static string urlBusqueda;
        private static string rama;
        private static string servicio;
        private static string nombreGrupo;
        private static string nombreTermSet;
        private static string field;
        private static string property;
        private static bool ramaVirgen;
        private static string campoOrden;
        private static StreamWriter logProceso;
        private static StreamWriter logDesprotegidos;
        private static TermStore termStore;
        private static TermSet termSet;
        private static bool grupos;
        private static bool desc;

        static void Main(string[] args)
        {
            CargaPropiedades(args);
            string timeStamp = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000") + ".txt";
            logProceso = new StreamWriter("LogProceso" + timeStamp);
            logDesprotegidos = new StreamWriter("LogDesprotegidos" + timeStamp);
            logDesprotegidos.WriteLine(rama);
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    TaxonomySession session = new TaxonomySession(site);
                    termStore = session.TermStores[servicio];
                    termSet = termStore.Groups[nombreGrupo].TermSets[nombreTermSet];
                    Term term = termSet.GetTerms(rama.Split('¬').Last(), false, StringMatchOption.ExactMatch, 100000, false).First(c => c.GetPath().ToLower().Replace(';','¬').Equals(rama.ToLower()));
                    if (grupos)
                    {
                        //if (ramaVirgen)
                        //{
                            AplicarGruposNuevos(web, term, term.Id.ToString());
                        //}
                        //else
                        //{
                        //    RefactorGrupos(web, term, term.Id.ToString());
                        //}
                    }
                    else
                    {
                        AplicarOrdenVerdes(web, term, term.Id.ToString());
                    }
                    logProceso.Close();
                    logDesprotegidos.Close();
                    StreamReader reader = new StreamReader("LogDesprotegidos" + timeStamp);
                    if (!reader.ReadToEnd().Contains("Documento Desprotegido"))
                    {
                        reader.Close();
                        File.Delete("LogDesprotegidos" + timeStamp);
                    }                    
                }
            }
        }

        private static void AplicarOrdenVerdes(SPWeb web, Term term, string guid)
        {
            ResultTable resultTable = RealizarBusqueda(guid);
            int contador = 0;
            List<DocumentoSharePoint> primerOrden = OrdenarResultados(resultTable);
            foreach (DocumentoSharePoint doc in primerOrden)
            {
                contador++;
                Console.WriteLine(doc.Path + " " + contador + "/" + primerOrden.Count);
                SPListItem item = web.GetFile(doc.Path).Item;
                if (item.File.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    TaxonomyField campo = (TaxonomyField)item.Fields.GetField(field);
                    int orden = contador * 100;
                    Console.WriteLine(orden);
                    item[campoOrden] = orden;
                    AplicarTerminoVerdes(term, doc.Path, item, orden);
                }
                else
                {
                    logDesprotegidos.WriteLine("Documento Desprotegido: " + doc.Path);
                }
            }
            if (ramaVirgen)
            {
                var borrables = term.Terms.Select(g => new { Guid = g.Id, Path = g.GetPath() });
                foreach (var borrar in borrables)
                {
                    try
                    {
                        Console.WriteLine("Se borrá: " + borrar.Path);
                        logProceso.WriteLine("Se borrá: " + borrar.Path);
                        termStore.GetTerm(borrar.Guid).Delete();
                        termStore.CommitAll();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error borrando el siguiente término: " + borrar.Path + " " + ex.Message);
                        logProceso.WriteLine("Error borrando el siguiente término: " + borrar.Path + " " + ex.Message);
                    }
                }
            }
        }
        //
        private static void AplicarTerminoVerdes(Term term, string path, SPListItem item, int orden)
        {
            try
            {
                TaxonomyField campo = (TaxonomyField)item.Fields.GetField(field);
                List<Term> terms = new List<Term>();
                ObtenerTerminosAplicados(term, path, item, term, terms);
                if (terms.Count > 0)
                {
                    campo.SetFieldValue(item, terms);
                }
                else
                {
                    campo.SetFieldValue(item, term);
                }
                item.SystemUpdate();
                logProceso.WriteLine(orden + " " + path + " " + term.GetPath());
            }
            catch (Exception ex)
            {
                logProceso.WriteLine("Error aplicando el metadato: " + path + " " + term.GetPath() + " " + ex.Message);
            }
        }

        private static void RefactorGrupos(SPWeb web, Term term, string guid)
        {
            ResultTable resultTable = RealizarBusqueda(guid);
            int contador = 0;
            int i = 0;
            List<DocumentoSharePoint> primerOrden = OrdenarResultados(resultTable);
            string[] guidsResoluciones = term.CustomSortOrder.Split(':').Reverse().ToArray();
            Guid guidResolucion = new Guid(guidsResoluciones[i]);
            Term resoulcion = termSet.GetTerm(guidResolucion);
            int prefijoGrupo = 10000;
            foreach (DocumentoSharePoint doc in primerOrden)
            {
                //contador++;
                contador += 100;
                //if (contador % 100 == 0)
                //{
                    //prefijoGrupo += 10000;
                    //contador = 1;
                    i++;
                    try
                    {
                        guidResolucion = new Guid(guidsResoluciones[i]);
                        resoulcion = termSet.GetTerm(guidResolucion);
                    }
                    catch (Exception ex) 
                    {
                        logProceso.WriteLine("Error obteniendo la nueva resolución: " + ex.Message + "se continuará aplicando en la rama: " + resoulcion.GetPath());
                    }
                //}
                Console.WriteLine(doc.Path + " " + contador + "/" + primerOrden.Count);
                SPListItem item = web.GetFile(doc.Path).Item;
                if (item.File.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    TaxonomyField campo = (TaxonomyField)item.Fields.GetField(field);
                    //int orden = prefijoGrupo + contador * 100;
                    //Console.WriteLine(orden);
                    Console.WriteLine(contador);
                    //item[campoOrden] = orden;
                    item[campoOrden] = contador;
                    AplicarTermino(term, resoulcion.GetPath().Split(';').Last(), doc.Path, item, resoulcion, contador);
                }
                else
                {
                    logDesprotegidos.WriteLine("Documento Desprotegido: " + doc.Path);
                }
            }
        }

        private static void AplicarGruposNuevos(SPWeb web, Term term, string guid)
        {
            ResultTable resultTable = RealizarBusqueda(guid);
            int contador = 0;
            List<DocumentoSharePoint> primerOrden = OrdenarResultados(resultTable);
            string grupo = "Grupo 1";
            int prefijoGrupo = 10000;
            int contGrupo = 1;
            string año = string.Empty;
            foreach (DocumentoSharePoint doc in primerOrden)
            {
                //contador++;
                contador += 100;
                if (!año.Equals("Año " + doc.Fecha.Year.ToString())) 
                {
                    año = "Año " + doc.Fecha.Year.ToString();
                }
                //if (contador % 100 == 0)
                //{
                //    prefijoGrupo += 10000;
                //    contGrupo++;
                //    contador = 1;
                //    grupo = "Grupo " + contGrupo;
                //}
                Console.WriteLine(doc.Path + " " + contador + "/" + primerOrden.Count);
                SPListItem item = web.GetFile(doc.Path).Item;
                if (item.File.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    Term tGrupo;
                    //tGrupo = ObtenerGrupo(term, grupo);
                    tGrupo = ObtenerGrupo(term, año);
                    //int orden = prefijoGrupo + contador * 100;
                    //Console.WriteLine(orden);
                    Console.WriteLine(contador);
                    //item[campoOrden] = orden;
                    item[campoOrden] = contador;
                    //AplicarTermino(term, grupo, doc.Path, item, tGrupo, orden);
                    AplicarTermino(term, año, doc.Path, item, tGrupo, contador);
                }
                else
                {
                    logDesprotegidos.WriteLine("Documento Desprotegido: " + doc.Path);
                }
            }
            if (ramaVirgen)
            {
                BorrarTerminos(term);
            }
            //string ordenTerminos = String.Join(":", term.Terms.Where(t => t.Name.StartsWith("Grupo")).OrderByDescending(t => t.Name).Select(t => t.Id.ToString()).ToArray());
            string ordenTerminos = String.Join(":", term.Terms.Where(t => t.Name.StartsWith("Año")).OrderByDescending(t => t.Name).Select(t => t.Id.ToString()).ToArray());
            term.CustomSortOrder = ordenTerminos;
            termStore.CommitAll();
        }

        private static Term ObtenerGrupo(Term term, string grupo)
        {
            Term tGrupo;
            try
            {
                tGrupo = term.Terms.First(t => t.Name.Equals(grupo));
            }
            catch (Exception ex)
            {
                tGrupo = term.CreateTerm(grupo, 3082);
                termStore.CommitAll();
                logProceso.WriteLine("Término creado: " + tGrupo.GetPath());
            }
            return tGrupo;
        }

        private static void AplicarTermino(Term term, string grupo, string path, SPListItem item, Term tGrupo, int orden)
        {
            try
            {
                TaxonomyField campo = (TaxonomyField)item.Fields.GetField(field);
                List<Term> terms = new List<Term>();
                ObtenerTerminosAplicados(term, path, item, tGrupo, terms);
                if (terms.Count > 0)
                {
                    campo.SetFieldValue(item, terms);
                }
                else
                {
                    campo.SetFieldValue(item, tGrupo);
                }
                item.SystemUpdate();
                logProceso.WriteLine(orden + " " + path + " " + term.GetPath() + ";" + grupo);
            }
            catch (Exception ex)
            {
                logProceso.WriteLine("Error aplicando el metadato: " + path + " " + term.GetPath() + ";" + grupo + " " + ex.Message);
            }
        }

        private static List<DocumentoSharePoint> OrdenarResultados(ResultTable resultTable)
        {
            List<DocumentoSharePoint> resultados = new List<DocumentoSharePoint>();
            List<string> primerOrden = new List<string>();
            DocumentoSharePoint dsp;
            while (resultTable.Read())
            {
                dsp = new DocumentoSharePoint();
                DateTime fechaBusqueda;
                bool esFecha = DateTime.TryParse(resultTable["Fecha"].ToString(),out fechaBusqueda);
                if (esFecha) 
                {
                    dsp.Fecha = fechaBusqueda;                    
                }
                dsp.Path = resultTable["Path"].ToString();
                dsp.Titulo = resultTable["Title"].ToString();
                resultados.Add(dsp);
            }
            if (desc) 
            {
                primerOrden = resultados.OrderBy(t => t.Path).Select(p => p.Path).ToList();
            }
            else
            {
                primerOrden = resultados.Where(t => t.Path.Split('/').Last().StartsWith("20110807")).OrderByDescending(t => t.Path).Select(p => p.Path).ToList().ToList();
                var segundoOrden = resultados.Where(t => !t.Path.Split('/').Last().StartsWith("20110807")).OrderBy(t => t).Select(p => p.Path).ToList();
                primerOrden.AddRange(segundoOrden);
            }
            return resultados.OrderBy(f => f.Fecha).ToList();
        }

        private static ResultTable RealizarBusqueda(string guid)
        {
            FullTextSqlQuery query = new FullTextSqlQuery(new SPSite(urlBusqueda));
            ResultType resultType = ResultType.RelevantResults;
            string strQuery = String.Format("SELECT Title, Path, Fecha FROM SCOPE() WHERE (\"" + property + "\"='#{0}')", guid);
            FullTextSqlQuery fullTextSqlQuery = new FullTextSqlQuery(new SPSite(urlBusqueda));
            fullTextSqlQuery.QueryText = strQuery;
            fullTextSqlQuery.ResultTypes = resultType;
            fullTextSqlQuery.RowLimit = 0;
            fullTextSqlQuery.TrimDuplicates = false;
            ResultTableCollection resultTableCollection = fullTextSqlQuery.Execute();
            ResultTable resultTable = resultTableCollection[resultType];
            return resultTable;
        }

        private static void BorrarTerminos(Term term)
        {
            //IEnumerable<Term> borrables = term.Terms.Where(t => !t.Name.StartsWith("Grupo"));
            IEnumerable<Term> borrables = term.Terms.Where(t => !t.Name.StartsWith("Año"));
            foreach (Term borrar in borrables)
            {
                string path = borrar.GetPath();
                try
                {
                    termSet.GetTerm(borrar.Id).Delete();
                    logProceso.WriteLine("Borrado: " + path);
                }
                catch (Exception ex)
                {
                    logProceso.WriteLine("Error tratando de borrar: " + path + " " + ex.Message);
                }
            }
        }

        private static void ObtenerTerminosAplicados(Term term, string path, SPListItem item, Term tGrupo, List<Term> terms)
        {
            string[] guids = item[field].ToString().Split(';').Where(t => !t.Split('|')[0].Equals(tGrupo.Name)).Select(t => t.Split('|')[1]).ToArray();
            if (guids.Count() > 0)
            {
                foreach (string g in guids)
                {
                    Term termActual = termSet.GetTerm(new Guid(g));
                    if (termActual != null && !termActual.GetPath().Replace(';', '¬').ToLower().Contains(rama.ToLower()))
                    {
                        logProceso.WriteLine(path + " posee: " + termActual.GetPath());
                        terms.Add(termActual);
                    }
                }
                terms.Add(tGrupo);
            }
        }

        private static void CargaPropiedades(string[] args)
        {
            AppSettingsReader app = new AppSettingsReader();
            url = (string)app.GetValue("Url", typeof(string));
            urlBusqueda = (string)app.GetValue("UrlBusqueda", typeof(string));
            rama = args[0];
            grupos = args[1].ToLower().Equals("grupos") ? true : false;
            ramaVirgen = args[2].ToLower().Equals("virgen") ? true : false;
            servicio = (string)app.GetValue("Servicio", typeof(string));
            nombreGrupo = (string)app.GetValue("Grupo", typeof(string));
            nombreTermSet = (string)app.GetValue("TermSet", typeof(string));
            field = (string)app.GetValue("Field", typeof(string));
            property = (string)app.GetValue("Property", typeof(string));
            campoOrden = (string)app.GetValue("CampoOrden", typeof(string));
            desc = args[3].ToLower().Equals("desc") ? true : false;
        }

    }
}
