using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Relativity.API;
using Relativity.Kepler.Transport;
using Relativity.Services.Interfaces.Document;
using Relativity.Services.Objects;
using Relativity.Services.Objects.DataContracts;

namespace RelativityAzureOcr
{
    [Description("Azure OCR")]
    [Guid("EAC80213-2702-4727-BE8A-F50940AC173D")]

    /*
     * Relativity Mass EventHandler Class
     */
    public class MassOperationHandler : kCura.MassOperationHandlers.MassOperationHandler
    {
        /*
         * Occurs after the user has selected items and pressed go.
         * In this function you can validate the items selected and return a warning/error message.
         */
        public override kCura.EventHandler.Response ValidateSelection()
        {
            // Get logger
            Relativity.API.IAPILog _logger = this.Helper.GetLoggerFactory().GetLogger().ForContext<MassOperationHandler>();

            // Init general response
            kCura.EventHandler.Response response = new kCura.EventHandler.Response()
            {
                Success = true,
                Message = ""
            };

            // Get current Workspace ID
            int workspaceId = this.Helper.GetActiveCaseID();
            _logger.LogDebug("Azure OCR, current Workspace ID: {workspaceId}", workspaceId.ToString());

            // Check if all Instance Settings are in place
            IDictionary<string, string> instanceSettings = this.GetInstanceSettings(ref response, new string[] { "DestinationField", "LogField", "AzureSubscriptionKey", "AzureEndpoint" });
            // Check if there was not error
            if (!response.Success)
            {
                return response;
            }

            // Count documents eligible for OCR
            // TODO: redo in Object Manager API
            try
            {
                // Construct and execute SQL Query to get the characters count
                string sqlText = "SELECT COUNT(*) FROM [EDDSDBO].[Document] AS [Document] JOIN [Resource].[" + this.MassActionTableName + "] AS [MassActionTableName] ON [Document].[ArtifactID] = [MassActionTableName].[ArtifactID] WHERE [Document].[HasNative] = 1";
                _logger.LogDebug("Azure OCR, document count SQL Parameter and Query: {query}", sqlText);
                long count = this.Helper.GetDBContext(workspaceId).ExecuteSqlStatementAsScalar<long>(sqlText);

                // Display number of documents to OCR
                _logger.LogDebug("Azure OCR, document count to OCR: {count}", count.ToString());
                response.Message = string.Format("Document count: {0}", count.ToString());
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, document count error");

                response.Success = false;
                response.Message = "Document count error";
            }

            return response;
        }

        /*
         * Occurs after the user has inputted data to a layout and pressed OK.
         * This function runs as a pre-save eventhandler.
         * This is NOT called if the mass operation does not have a layout.
         */
        public override kCura.EventHandler.Response ValidateLayout()
        {
            // Init general response
            kCura.EventHandler.Response response = new kCura.EventHandler.Response()
            {
                Success = true,
                Message = ""
            };
            return response;
        }

        /*
         * Occurs before batching begins. A sample use would be to set up an instance of an object.
         */
        public override kCura.EventHandler.Response PreMassOperation()
        {
            // Init general response
            kCura.EventHandler.Response response = new kCura.EventHandler.Response()
            {
                Success = true,
                Message = ""
            };
            return response;
        }

        /*
         * This function is called in batches based on the size defined in configuration.
         */
        public override kCura.EventHandler.Response DoBatch()
        {
            // Get logger
            Relativity.API.IAPILog _logger = this.Helper.GetLoggerFactory().GetLogger().ForContext<MassOperationHandler>();

            // Init general response
            kCura.EventHandler.Response response = new kCura.EventHandler.Response()
            {
                Success = true,
                Message = ""
            };

            // Get current Workspace ID
            int workspaceId = this.Helper.GetActiveCaseID();
            _logger.LogDebug("Azure OCR, current Workspace ID: {workspaceId}", workspaceId.ToString());

            // Check if all Instance Settings are in place
            IDictionary<string, string> instanceSettings = this.GetInstanceSettings(ref response, new string[] { "DestinationField", "LogField", "AzureSubscriptionKey", "AzureEndpoint"});
            // Check if there was not error
            if (!response.Success)
            {
                return response;
            }

            // Update general status
            this.ChangeStatus("OCRing documents");

            // For each document create OCR task
            List<Task<int>> ocrTasks = new List<Task<int>>();
            int runningTasks = 0;
            int concurrentTasks = 16;
            for (int i = 0; i < this.BatchIDs.Count; i++)
            {
                // OCR documents in Azure and update Relativity using Object Manager API
                ocrTasks.Add(OcrDocument(workspaceId, this.BatchIDs[i], instanceSettings["DestinationField"], instanceSettings["LogField"], instanceSettings["AzureSubscriptionKey"], instanceSettings["AzureEndpoint"]));

                // Update progreass bar
                this.IncrementCount(1);

                // Allow only certain number of tasks to run concurrently
                do
                {
                    runningTasks = 0;
                    foreach (Task<int> ocrTask in ocrTasks)
                    {
                        if (!ocrTask.IsCompleted)
                        {
                            runningTasks++;
                        }
                    }
                    if (runningTasks >= concurrentTasks)
                    {
                        Thread.Sleep(100);
                    }
                } while (runningTasks >= concurrentTasks);
            }

            // Update general status
            this.ChangeStatus("Waiting to finish the document OCR");

            // Wait for all OCRs to finish
            _logger.LogDebug("Azure OCR, waiting for all documents finish OCRing ({n} document(s))", this.BatchIDs.Count.ToString());
            Task.WaitAll(ocrTasks.ToArray());

            // Update general status
            this.ChangeStatus("Checking the results of the document OCR");

            // Check results
            List<string> ocrErrors = new List<string>();
            for (int i = 0; i < ocrTasks.Count; i++)
            {
                // If OCR was not done add to the error List
                _logger.LogDebug("Azure OCR, OCR task result: {result} (task: {task})", ocrTasks[i].Result.ToString(), ocrTasks[i].Id.ToString());
                if (ocrTasks[i].Result != 0)
                {
                    ocrErrors.Add(ocrTasks[i].Result.ToString());
                }
            }

            // If there are any errors adjust response
            if (ocrErrors.Count > 0)
            {
                _logger.LogError("Azure OCR, not all documents have been OCRed: ({documents})", string.Join(", ", ocrErrors));

                response.Success = false;
                response.Message = "Not all documents have been OCRed";
            }

            return response;
        }

        /*
         * Occurs after all batching is completed.
         */
        public override kCura.EventHandler.Response PostMassOperation()
        {
            // Init general response
            kCura.EventHandler.Response response = new kCura.EventHandler.Response()
            {
                Success = true,
                Message = ""
            };
            return response;
        }

        /*
         * Custom method to get required Relativity Instance Settings
         */
        private IDictionary<string, string> GetInstanceSettings(ref kCura.EventHandler.Response response, string[] instanceSettingsNames)
        {
            // Get logger
            Relativity.API.IAPILog _logger = this.Helper.GetLoggerFactory().GetLogger().ForContext<MassOperationHandler>();

            // Output Dictionary
            IDictionary<string, string> instanceSettingsValues = new Dictionary<string, string>();

            // Get and validate instance settings
            foreach (string name in instanceSettingsNames)
            {
                try
                {
                    instanceSettingsValues.Add(name, this.Helper.GetInstanceSettingBundle().GetString("Azure.OCR", name));
                    if (instanceSettingsValues[name].Length <= 0)
                    {
                        _logger.LogError("Azure OCR, Instance Settings empty error: {section}/{name}", "Azure.OCR", name);

                        response.Success = false;
                        response.Message = "Instance Settings error";
                        return instanceSettingsValues;
                    }
                }
                catch (Exception e)
                {
                    _logger.LogError(e, "Azure OCR, Instance Settings error: {section}/{name}", "Azure.OCR", name);

                    response.Success = false;
                    response.Message = "Instance Settings error";
                    return instanceSettingsValues;
                }

                _logger.LogDebug("Azure OCR, Instance Setting: {name}=>{value}", name, instanceSettingsValues[name]);
            }

            return instanceSettingsValues;
        }

        /*
         * Custom method to OCR document using Azure Computer Vision
         */
        private async Task<int> OcrDocument(int workspaceId, int documentArtifactId, string destinationField, string logField, string azureSubscriptionKey, string azureEndpoint)
        {
            // Get logger
            Relativity.API.IAPILog _logger = this.Helper.GetLoggerFactory().GetLogger().ForContext<MassOperationHandler>();

            // Get Relativity Document File Manager API
            IDocumentFileManager documentFileManager = this.Helper.GetServicesManager().CreateProxy<IDocumentFileManager>(ExecutionIdentity.CurrentUser);

            // Get Relativity Object Manager API
            IObjectManager objectManager = this.Helper.GetServicesManager().CreateProxy<IObjectManager>(ExecutionIdentity.CurrentUser);

            // Check if document has native
            try
            {
                // Construct objects and retreive document content
                ReadRequest readRequest = new ReadRequest
                {
                    Object = new RelativityObjectRef
                    {
                        ArtifactID = documentArtifactId
                    },
                    Fields = new List<FieldRef>()
                    {
                        new FieldRef
                        {
                            Name = "Has Native"
                        }
                    }
                };
                ReadResult readResult = await objectManager.ReadAsync(workspaceId, readRequest);
                if (!(bool)readResult.Object.FieldValues[0].Value)
                {
                    // Return 0 as there is no native
                    _logger.LogDebug("Azure OCR, document does not have native (ArtifactID: {id})", documentArtifactId.ToString());
                    return 0;
                }
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, native check of the document for OCR error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }

            // Get document native
            Stream streamToOcr;
            try
            {
                // Construct objects and retreive native document
                IKeplerStream keplerStream = await documentFileManager.DownloadNativeFileAsync(workspaceId, documentArtifactId);
                streamToOcr = await keplerStream.GetStreamAsync();
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, native of the document for OCR retrieval error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }

            // Log original native document
            _logger.LogDebug("Azure OCR, original native document downloaded (ArtifactID: {id})", documentArtifactId.ToString());

            // Force TLS 1.2 or higher as Azure requires it
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 & ~(SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11);

            // Build OCR analyze request
            HttpRequestMessage requestAnalyze = new HttpRequestMessage();
            requestAnalyze.Method = HttpMethod.Post;
            requestAnalyze.RequestUri = new Uri(azureEndpoint + "vision/v3.2/read/analyze"); // https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/quickstarts-sdk/client-library?tabs=visual-studio&pivots=programming-language-rest-api
            requestAnalyze.Content = new StreamContent(streamToOcr);
            requestAnalyze.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            requestAnalyze.Headers.Add("Ocp-Apim-Subscription-Key", azureSubscriptionKey);

            // Send the analyze request
            HttpClient clientAnalyze = new HttpClient();
            HttpResponseMessage responseAnalyze = await clientAnalyze.SendAsync(requestAnalyze).ConfigureAwait(false);

            // Check the response
            if (!responseAnalyze.IsSuccessStatusCode)
            {
                _logger.LogError("Azure OCR, HTTP analyze reposnse error (ArtifactID: {id}, status: {status})", documentArtifactId.ToString(), responseAnalyze.StatusCode.ToString());
                return documentArtifactId;
            }

            // Read Operation-Location header from the analyze response
            string operationLocation = "";
            try
            {
                operationLocation = responseAnalyze.Headers.GetValues("Operation-Location").First();
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, Operation-Location header retrieval error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }
            clientAnalyze.Dispose();

            // Log Operation-Location header from the response
            _logger.LogDebug("Azure OCR, Operation-Location header: {header})", operationLocation);

            // Send the request, it may not be available right away so keep trying
            string ocredStatus = "";
            string ocred = "";
            int retries = 50; // 50 == 5min
            do
            {
                --retries;
                await Task.Delay(100);

                // Build OCR read request
                HttpRequestMessage requestRead = new HttpRequestMessage();
                requestRead.Method = HttpMethod.Get;
                requestRead.RequestUri = new Uri(operationLocation);
                requestRead.Headers.Add("Ocp-Apim-Subscription-Key", azureSubscriptionKey);

                // Send the read request
                HttpClient clientRead = new HttpClient();
                HttpResponseMessage responseRead = await clientRead.SendAsync(requestRead).ConfigureAwait(false);

                // Check the response
                if (!responseRead.IsSuccessStatusCode)
                {
                    _logger.LogError("Azure OCR, HTTP read reposnse error (ArtifactID: {id}, status: {status})", documentArtifactId.ToString(), responseRead.StatusCode.ToString());
                    return documentArtifactId;
                }

                // Read the read response
                ocred = await responseRead.Content.ReadAsStringAsync();
                _logger.LogDebug("Azure OCR, OCR result (ArtifactID: {id}, length: {length}, result:{result})", documentArtifactId.ToString(), ocred.Length.ToString(), ocred);
                clientRead.Dispose();

                // Parse JSON and get the OCR result status
                try
                {
                    JsonElement ocrResults = JsonSerializer.Deserialize<JsonElement>(ocred);

                    // Check the OCR result
                    ocredStatus = ocrResults.GetProperty("status").GetString();
                    _logger.LogDebug("Azure OCR, OCR result status (ArtifactID: {id}, OCR status:{status})", documentArtifactId.ToString(), ocredStatus);
                }
                catch (Exception e)
                {
                    _logger.LogError(e, "Azure OCR, OCR result status JSON read error (ArtifactID: {id})", documentArtifactId.ToString());
                    return documentArtifactId;
                }
            } while (retries >= 0 && ocredStatus != "succeeded");

            if (retries < 0)
            {
                _logger.LogError("Azure OCR, OCR was not fininshed within the timeout (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }

            // Parse JSON and get the OCR text
            List<string> linesOcred = new List<string>();
            try
            {
                JsonElement ocrResults = JsonSerializer.Deserialize<JsonElement>(ocred);

                // Get all OCRed lines
                foreach (JsonElement readResults in ocrResults.GetProperty("analyzeResult").GetProperty("readResults").EnumerateArray())
                {
                    foreach (JsonElement lines in readResults.GetProperty("lines").EnumerateArray())
                    {
                        linesOcred.Add(lines.GetProperty("text").GetString());
                    }
                }

                // Log the OCR result
                _logger.LogDebug("Azure OCR, OCR check (ArtifactID: {id}, no. lines: {length})", documentArtifactId.ToString(), linesOcred.Count.ToString());
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, OCR result JSON read error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }

            // Construct OCRed text
            string textOcred = string.Join(string.Empty, linesOcred);
            Stream streamOcred = new MemoryStream();
            StreamWriter streamWriter = new StreamWriter(streamOcred);
            streamWriter.Write(textOcred);
            streamWriter.Flush();
            streamOcred.Position = 0;

            // Log OCRed document
            _logger.LogDebug("Azure OCR, OCRed document (ArtifactID: {id}, length: {length})", documentArtifactId.ToString(), textOcred.Length.ToString());

            // Update document OCRed text
            try
            {
                // Construct objects and do document update
                RelativityObjectRef relativityObject = new RelativityObjectRef
                {
                    ArtifactID = documentArtifactId
                };
                FieldRef relativityField = new FieldRef
                {
                    Name = destinationField
                };
                UpdateLongTextFromStreamRequest updateRequest = new UpdateLongTextFromStreamRequest
                {
                    Object = relativityObject,
                    Field = relativityField
                };
                KeplerStream keplerStream = new KeplerStream(streamOcred);
                await objectManager.UpdateLongTextFromStreamAsync(workspaceId, updateRequest, keplerStream);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, document for OCR update error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }
            
            // Update document OCR log
            try
            {
                Stream streamCurrentLog;
                // Construct objects and get current OCR log
                RelativityObjectRef relativityObject = new RelativityObjectRef
                {
                    ArtifactID = documentArtifactId
                };
                FieldRef relativityField = new FieldRef
                {
                    Name = logField
                };
                IKeplerStream keplerStream = await objectManager.StreamLongTextAsync(workspaceId, relativityObject, relativityField);
                streamCurrentLog = await keplerStream.GetStreamAsync();

                // Add new OCR log
                Stream streamUpdatedLog = new MemoryStream();
                StreamWriter streamLogWriter = new StreamWriter(streamUpdatedLog);
                streamLogWriter.Write(new StreamReader(streamCurrentLog).ReadToEnd());
                streamLogWriter.Write("Azure OCR;" + this.Helper.GetAuthenticationManager().UserInfo.EmailAddress + ";" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + ";" + textOcred.Length.ToString() + "\n");
                streamLogWriter.Flush();
                streamUpdatedLog.Position = 0;

                // Write updated OCR log
                UpdateLongTextFromStreamRequest updateRequest = new UpdateLongTextFromStreamRequest
                {
                    Object = relativityObject,
                    Field = relativityField
                };
                keplerStream = new KeplerStream(streamUpdatedLog);
                await objectManager.UpdateLongTextFromStreamAsync(workspaceId, updateRequest, keplerStream);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Azure OCR, OCR log update error (ArtifactID: {id})", documentArtifactId.ToString());
                return documentArtifactId;
            }
            
            // Return 0 as all went without error
            return 0;
        }
    }
}