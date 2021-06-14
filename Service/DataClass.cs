
using Mustache;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using WebApplication2.Model;
using HtmlAgilityPack;
using System.Text;
using System.Reflection;

namespace WebApplication2.Service
{
    public class DataClass
    {
        public string GetDecode()
        {

            return "";
        }
        public Dictionary<string, string> FillReport(ModelTest param)
        {
            Console.WriteLine("NIDALERTA: {0}, NPERIODO_PROCESO: {1}, NIDUSUARIO_ASIGNADO: {2}", param.NIDALERTA, param.NPERIODO_PROCESO, param.NIDUSUARIO_ASIGNADO);

            var id = Guid.NewGuid();
            string ruta = string.Format("C:/plantillasLaft/{0}/{1}/", param.NIDALERTA, id);
            Environment.ExpandEnvironmentVariables(ruta);
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            var rutaTemplate = string.Format("{0}/{1}/{2}.docx", "C:/plantillasLaft", param.NIDALERTA, param.SNOMBRE_ALERTA);
            var rutaUsuario = string.Format("{0}/{1}.docx", ruta, param.SNOMBRE_ALERTA);
            var rutaExtract = string.Format("{0}/{1}", ruta, "ex");
            File.Copy(rutaTemplate, rutaUsuario);
            Directory.CreateDirectory(rutaExtract);
            ZipFile.ExtractToDirectory(rutaUsuario, rutaExtract);
            var docExtract = string.Format("{0}/{1}", rutaExtract, "word/document.xml");
            var texto = File.ReadAllText(docExtract);
            List<Dictionary<string, dynamic>> lista = this.getData();
            Console.WriteLine("lista: {0}", lista.Count);
            lista.ForEach(it => Console.WriteLine("it: {0} param: {1} et: {2} rpt: {3}", it["NIDUSUARIO_ASIGNADO"], param.NIDUSUARIO_ASIGNADO, it["ETIQUETA"], it["SRESPUESTA"]));
            List<Dictionary<string, dynamic>> listaPorUsuario = lista.Where(it => param.NIDUSUARIO_ASIGNADO.ToString().Equals(it["NIDUSUARIO_ASIGNADO"].ToString()) || it["SOBLIGA_USUARIO"].ToString().Equals("1")).ToList();
            //List<Dictionary<string, dynamic>> listaPorUsuario = lista.ToList ();
            Console.WriteLine("listaPorUsuario : " + listaPorUsuario);
            Console.WriteLine("listaPorUsuario length: " + listaPorUsuario.Count);
            var compiler = new FormatCompiler();
            var generator = compiler.Compile(texto);
            var resultado = this.generarTemplate(listaPorUsuario, generator);
            File.Delete(docExtract);
            File.Delete(rutaUsuario);
            File.WriteAllText(docExtract, resultado);
            ZipFile.CreateFromDirectory(rutaExtract, string.Format("{0}/{1}-nuevo.pdf", ruta, param.SNOMBRE_ALERTA));






            byte[] base64 = File.ReadAllBytes(string.Format("{0}/{1}-nuevo.pdf", ruta, param.SNOMBRE_ALERTA));
            string base64String = Convert.ToBase64String(base64);
            var output = new Dictionary<string, string>();
            output["base64"] = base64String;
            return output;
        }

        public Dictionary<string, string> GetPdf(ModelTest param)
        {
            Dictionary<string, string> output = new Dictionary<string, string>();

            var id = Guid.NewGuid();
            string ruta = string.Format("C:/plantillasLaft/{0}/{1}/", param.NIDALERTA, id);
            Environment.ExpandEnvironmentVariables(ruta);
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            var rutaTemplate = string.Format("{0}/{1}/{2}.html", "C:/plantillasLaft", param.NIDALERTA, param.SNOMBRE_ALERTA);
            var rutaUsuario = string.Format("{0}/{1}.html", ruta, param.SNOMBRE_ALERTA);
            var rutaExtract = string.Format("{0}/{1}", ruta, "ex");
            File.Copy(rutaTemplate, rutaUsuario);
            Directory.CreateDirectory(rutaExtract);
            var doc = new HtmlDocument();
            doc.Load(rutaUsuario);

            List<Dictionary<string, dynamic>> lista = this.getData();
            //List<Dictionary<string, dynamic>> listaPorUsuario = lista.Where(it => param.NIDUSUARIO_ASIGNADO.ToString().Equals(it["NIDUSUARIO_ASIGNADO"].ToString()) || it["SOBLIGA_USUARIO"].ToString().Equals("1")).ToList();
            File.Delete(rutaUsuario);
            HtmlDocument resultado = getDocResult(lista, doc);
            resultado.Save(string.Format("{0}/{1}-nuevo.html", ruta, param.SNOMBRE_ALERTA), Encoding.UTF8);


            string path = AppDomain.CurrentDomain.BaseDirectory + "apps/wkhtmltopdf.exe";
            path = path.Replace("\\bin\\Debug\\netcoreapp3.1", "");
            ProcessStartInfo nprocessS = new ProcessStartInfo();
            nprocessS.UseShellExecute = false;
            nprocessS.FileName = path;
            nprocessS.Arguments = $"{ruta}/{param.SNOMBRE_ALERTA}-nuevo.html {ruta}/{param.SNOMBRE_ALERTA}-nuevo.pdf";


            using (Process p = Process.Start(nprocessS))
            {
                p.WaitForExit();
            }

            byte[] base64 = File.ReadAllBytes(string.Format("{0}/{1}-nuevo.pdf", ruta, param.SNOMBRE_ALERTA));
            string base64String = Convert.ToBase64String(base64);
            output["base64"] = base64String;

            //output = new Dictionary<string, string>();
            return output;
        }

        private HtmlDocument getDocResult(List<Dictionary<string, dynamic>> lista, HtmlDocument doc)
        {
            String html = doc.ParsedText;
            for (int i = 0; i < lista.Count; i++)
            {
                try
                {
                    if (lista[i]["ETIQUETA"].StartsWith("tbl"))
                    {
                   
                    }
                    else
                    {
                        html = html.Replace($"${lista[i]["ETIQUETA"]}", lista[i]["SRESPUESTA"]);
                    }
                }
                catch (Exception ex) { throw; }
            }
            doc.LoadHtml(html);
            for (int i = 0; i < lista.Count; i++)
            {
                try
                {
                    if (lista[i]["ETIQUETA"].StartsWith("tbl"))
                    {
                        List<Tabla> listaTabla = this.fillTable(lista[i]);
                        HtmlNode docNodeTBody = doc.GetElementbyId(lista[i]["ETIQUETA"]);
                        if (listaTabla.Count > 0)
                        {
                            for (int l = 0; l < listaTabla.Count; l++)
                            {
                                TypeInfo typeInfo = typeof(Tabla).GetTypeInfo();
                                var props = typeInfo.DeclaredProperties;
                                HtmlNode docNodeTr = new HtmlDocument().CreateElement("tr");
                                foreach (var prop in props)
                                {
                                    string sValue = (prop.GetValue(listaTabla[l]) ?? "").ToString();
                                    if (sValue != "")
                                    {
                                        HtmlNode docNodetd = new HtmlDocument().CreateElement("td");
                                        docNodetd.InnerHtml = sValue;
                                        docNodeTr.ChildNodes.Append(docNodetd);
                                    }
                                }
                                docNodeTBody.ChildNodes.Append(docNodeTr);
                            }
                        }
                    }
                }
                catch (Exception ex) { throw; }
            }
            return doc;
        }

        private List<Dictionary<string, dynamic>> getData()
        {
            List<Dictionary<string, dynamic>> items = new List<Dictionary<string, dynamic>>();
            Dictionary<string, dynamic> item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "DNI|08243490|ALIAGA ABANTO CARLOS ANIBAL|RT|DNI|08429294|ALIAGA SUAREZ DE LAU ROSA MARIA|RT");
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "tblResultados");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", "0");
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "2020");
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "periodoAnio");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", "0");
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "09");
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "periodoMes");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", "0");
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "2");
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "nCantidadResultados");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", "0");
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "0");
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "nCantidadClienteNegativo");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", "0");
            items.Add(item);
            return items;
        }
        private List<Dictionary<string, dynamic>> getData1()
        {
            List<Dictionary<string, dynamic>> items = new List<Dictionary<string, dynamic>>();
            Dictionary<string, dynamic> item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "NO");
            item.Add("NIDUSUARIO_ASIGNADO", 72);
            item.Add("ETIQUETA", "rpta712");
            item.Add("NIDREGIMEN", 1);
            item.Add("SOBLIGA_USUARIO", 0);
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "Si");
            item.Add("NIDUSUARIO_ASIGNADO", 2);
            item.Add("ETIQUETA", "rpta711");
            item.Add("NIDREGIMEN", 1);
            item.Add("SOBLIGA_USUARIO", 0);
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "Si");
            item.Add("NIDUSUARIO_ASIGNADO", 2);
            item.Add("ETIQUETA", "rpta712");
            item.Add("NIDREGIMEN", 1);
            item.Add("SOBLIGA_USUARIO", 0);
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", "Si");
            item.Add("NIDUSUARIO_ASIGNADO", 72);
            item.Add("ETIQUETA", "rpta711");
            item.Add("NIDREGIMEN", 1);
            item.Add("SOBLIGA_USUARIO", 0);
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", 2020);
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "rpta713");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", 1);
            items.Add(item);
            item = new Dictionary<string, dynamic>();
            item.Add("SRESPUESTA", 09);
            item.Add("NIDUSUARIO_ASIGNADO", 0);
            item.Add("ETIQUETA", "periodoMes");
            item.Add("NIDREGIMEN", 0);
            item.Add("SOBLIGA_USUARIO", 1);
            items.Add(item);
            return items;
        }
        private string generarTemplate(List<Dictionary<string, dynamic>> listaPorUsuario, dynamic generator)
        {
            try
            {
                Console.WriteLine("generar : " + listaPorUsuario);
                dynamic valores = new ExpandoObject();
                foreach (var item in listaPorUsuario)
                {
                    Console.WriteLine("generar item : " + item);
                    try
                    {
                        Console.WriteLine("el item : " + item);
                        if (item["ETIQUETA"].StartsWith("tbl"))
                        {
                            List<Tabla> listaTabla = this.fillTable(item);
                            ((IDictionary<string, object>)valores).Add(item["ETIQUETA"], listaTabla);
                        }
                        else
                        {
                            ((IDictionary<string, object>)valores).Add(item["ETIQUETA"], item["SRESPUESTA"]);
                        }
                    }
                    catch (Exception ex) { throw; }
                }
                Console.WriteLine("el valores : " + valores);
                string resultado = generator.Render(valores);
                return resultado;
            }
            catch (Exception ex)
            {
                Console.WriteLine("el error : " + ex);
                throw;
            }

        }
        private List<Tabla> fillTable(Dictionary<string, dynamic> item)
        {
            try
            {
                List<Tabla> listaTabla = new List<Tabla>();
                var valores = item["SRESPUESTA"].Split("|");
                var listaValores = new List<string>(valores);
                while (listaValores.Any())
                {
                    var tabla = new Tabla();
                    var tablaType = tabla.GetType();
                    for (int i = 0; i < 4; i++)
                    {
                        var primerElemento = listaValores[0];
                        var k = i + 1;
                        tablaType.GetProperty("valor" + k).SetValue(tabla, primerElemento, null);
                        listaValores.RemoveAt(0);

                    }
                    Console.WriteLine("el tale : " + tabla);
                    Console.WriteLine(tabla);
                    listaTabla.Add(tabla);
                }

                Console.WriteLine(listaTabla);
                return listaTabla;
            }
            catch (Exception ex)
            {
                Console.WriteLine("el error en el fillTable : " + ex);
                throw;
            }
        }
        public class Tabla
        {
            public string valor1 { get; set; }
            public string valor2 { get; set; }
            public string valor3 { get; set; }
            public string valor4 { get; set; }
            public string valor5 { get; set; }
            public string valor6 { get; set; }
        }
    }
}
