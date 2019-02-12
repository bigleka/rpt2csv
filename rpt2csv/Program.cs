using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpt2csv
{
    class Program
    {
        /// <summary>
        /// </summary>
        /// <param name="args">argumentos.</param>
        internal static void Main(string[] args)
        {
            try //para o tratamento de exception em caso de arquivo nao encontrado
            {
                if (args.Contains("/?"))
                {
                    Console.WriteLine("Converte o arquivo de saída .RPT do SQL Server Management Studio em um arquivo .CSV .");
                    Console.WriteLine("Você pode gerar um arquivo .RPT selecionando \"Results to File\" na barra de ferramentas.");
                    Console.WriteLine();
                    Console.WriteLine("Utilização: Rpt2Csv.exe <arquivo1> [<arquivo2> ...]");
                    Console.WriteLine();
                    Console.WriteLine("Exemplo: Rpt2Csv.exe ResultadoTeste.RPT");
                    Console.WriteLine();
                    Console.WriteLine("No mesmo local onde está o arquivo ResultadoTeste.RPT será criado um arquivo ResultadoTeste.CSV");
                    Console.WriteLine("Caso o arquivo .RPT possua mais de um resultado de query apenas o primerio resultado estará no arquivo .CSV");
                    return;
                }
                if (args.Length > 0)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        string inputFile;
                        string outputFile;

                        inputFile = args[i];
                        outputFile = Path.GetFileNameWithoutExtension(args[i]) + ".csv";

                        Environment.CurrentDirectory = Path.GetDirectoryName(inputFile).Length == 0 ? Environment.CurrentDirectory : Path.GetFullPath(Path.GetDirectoryName(inputFile));

                        using (StreamReader inputReader = File.OpenText(inputFile))
                        {
                            string firstLine = inputReader.ReadLine();
                            string secondLine = inputReader.ReadLine();

                            string[] underscores = secondLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                            string[] fields = new string[underscores.Length];
                            int[] fieldLengths = new int[underscores.Length];

                            for (int j = 0; j < fieldLengths.Length; j++)
                            {
                                fieldLengths[j] = underscores[j].Length;
                            }

                            int fileNumber = 0;

                            StreamWriter outputWriter = null;

                            try
                            {
                                outputWriter = File.CreateText(outputFile.Insert(outputFile.LastIndexOf("."), "_" + fileNumber.ToString()));
                                fileNumber++;

                                int lineNumber = 0;

                                WriteLineToCsv(outputWriter, fieldLengths, firstLine);
                                lineNumber++;

                                string line;

                                while ((line = inputReader.ReadLine()) != null)
                                {
                                    //if (lineNumber >= 65536)
                                    //{
                                    //    outputWriter.Close();
                                    //    outputWriter = File.CreateText(outputFile.Insert(outputFile.LastIndexOf("."), "_" + fileNumber.ToString()));
                                    //    fileNumber++;

                                    //    lineNumber = 0;

                                    //    WriteLineToCsv(outputWriter, fieldLengths, firstLine);
                                    //    lineNumber++;
                                    //}

                                    if (!WriteLineToCsv(outputWriter, fieldLengths, line))
                                    {
                                        break;
                                    }

                                    lineNumber++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex);
                                Console.WriteLine("NOTA: O arquivo de entrada não pode conter qualquer caractere de entrada de nova linha como coluna.");
                                Console.WriteLine();
                                Console.WriteLine("Pressione qualquer tecla para continuar...");
                                Console.ReadKey(true);
                            }
                            finally
                            {
                                if (outputWriter != null)
                                {
                                    outputWriter.Close();
                                }
                            }

                            // Se for apenas um arquivo, não precisamos do número no nome do arquivo.
                            if (fileNumber == 1)
                            {
                                try
                                {
                                    if (File.Exists(outputFile))
                                    {
                                        File.Delete(outputFile);
                                    }

                                    File.Move(outputFile.Insert(outputFile.LastIndexOf("."), "_0"), outputFile);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex);
                                    Console.WriteLine("Press any key to continue...");
                                    Console.ReadKey(true);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Converte o arquivo de saída .RPT do SQL Server Management Studio em um arquivo .CSV .");
                    Console.WriteLine("Você pode gerar um arquivo .RPT selecionando \"Results to File\" na barra de ferramentas.");
                    Console.WriteLine();
                    Console.WriteLine("Utilização: Rpt2Csv.exe <arquivo1> [<arquivo2> ...]");
                    return;
                }
            }
            catch //o try da linha 18
            {
                Console.WriteLine("Alguma coisa aconteceu de errado.");
                Console.WriteLine("Ou o arquivo .RPT não existe ou você está tentando adicionar o arquivo de destino .CSV como parâmetro de execução.");
                Console.WriteLine();
                Console.WriteLine("Utilização: Rpt2Csv.exe <arquivo1> [<arquivo2> ...]");
                return;
            }
        }

        /// <summary>
        /// Converte uma única linha de campo de largura fixa para uma única linha de campo separado por vírgula.
        /// </summary>
        /// <param name="outputWriter">O caractere usado para a separação das colunas.</param>
        /// <param name="fieldLengths">Uma matriz contendo os comprimentos dos campos.</param>
        /// <param name="line">A linha que será convertida para CSV.</param>
        /// <returns>Verdadeiro se conseguir converter a linha, caso contrario Falso.</returns>
        private static bool WriteLineToCsv(StreamWriter outputWriter, int[] fieldLengths, string line)
        {
            if (line.Length == 0)
            {
                return false;
            }

            int index = 0;

            for (int i = 0; i < fieldLengths.Length; i++)
            {
                string value;

                if (i < fieldLengths.Length - 1)
                {
                    value = line.Substring(index, fieldLengths[i]);
                }
                else
                {
                    value = line.Substring(index);
                }

                value = value.Replace("\"", "\"\"");
                value = value.Trim();

                if (value == "NULL")
                {
                    value = string.Empty;
                }

                outputWriter.Write("\"{0}\"", value);
                index += fieldLengths[i] + 1;

                if (i < fieldLengths.Length - 1)
                {
                    outputWriter.Write(",");
                }
                else
                {
                    outputWriter.WriteLine();
                }
            }

            return true;
        }
    }
}
