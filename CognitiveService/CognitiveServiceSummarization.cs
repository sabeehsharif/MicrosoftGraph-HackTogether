using Azure;
using System;
using Azure.AI.TextAnalytics;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Graph;
using System.Collections;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.ExternalConnectors;

namespace DotNetCoreRazor_MSGraph.CognitiveService
{
    //namespace MicrosoftGraph-DotNetCoreRazor.CognitiveService
    //DotNetCoreRazor_MSGraph.Graph
    public class CognitiveServiceSummarization
{
        private static AzureKeyCredential credentials;
        private static Uri endpoint;
        private static List<string> summarizedText = new List<string>();
    // Example method for summarizing text
    public static async Task TextSummarizationExample(TextAnalyticsClient client, string document)
    {
        //string document = @"The extractive summarization feature uses natural language processing techniques to locate key sentences in an unstructured text document. 
        //        These sentences collectively convey the main idea of the document. This feature is provided as an API for developers. 
        //        They can use it to build intelligent solutions based on the relevant information extracted to support various use cases. 
        //        In the public preview, extractive summarization supports several languages. It is based on pretrained multilingual transformer models, part of our quest for holistic representations. 
        //        It draws its strength from transfer learning across monolingual and harness the shared nature of languages to produce models of improved quality and efficiency.";

        // Prepare analyze operation input. You can add multiple documents to this list and perform the same
        // operation to all of them.
        var batchInput = new List<string>
            {
                document
            };

        TextAnalyticsActions actions = new TextAnalyticsActions()
        {
            ExtractSummaryActions = new List<ExtractSummaryAction>() { new ExtractSummaryAction() }
        };

        // Start analysis process.
        AnalyzeActionsOperation operation = await client.StartAnalyzeActionsAsync(batchInput, actions);
        await operation.WaitForCompletionAsync();
        // View operation status.
        //Console.WriteLine($"AnalyzeActions operation has completed");
        //Console.WriteLine();

        //Console.WriteLine($"Created On   : {operation.CreatedOn}");
        //Console.WriteLine($"Expires On   : {operation.ExpiresOn}");
        //Console.WriteLine($"Id           : {operation.Id}");
        //Console.WriteLine($"Status       : {operation.Status}");

        //Console.WriteLine();
        // View operation results.
        await foreach (AnalyzeActionsResult documentsInPage in operation.Value)
        {
                summarizedText.Clear();

                IReadOnlyCollection<ExtractSummaryActionResult> summaryResults = documentsInPage.ExtractSummaryResults;

            foreach (ExtractSummaryActionResult summaryActionResults in summaryResults)
            {
                //if (summaryActionResults.HasError)
                //{
                //    Console.WriteLine($"  Error!");
                //    Console.WriteLine($"  Action error code: {summaryActionResults.Error.ErrorCode}.");
                //    Console.WriteLine($"  Message: {summaryActionResults.Error.Message}");
                //    continue;
                //}

                foreach (ExtractSummaryResult documentResults in summaryActionResults.DocumentsResults)
                {
                        //if (documentResults.HasError)
                        //{
                        //    Console.WriteLine($"  Error!");
                        //    Console.WriteLine($"  Document error code: {documentResults.Error.ErrorCode}.");
                        //    Console.WriteLine($"  Message: {documentResults.Error.Message}");
                        //    continue;
                        //}

                        //Console.WriteLine($"  Extracted the following {documentResults.Sentences.Count} sentence(s):");
                        //Console.WriteLine();
                    foreach (SummarySentence sentence in documentResults.Sentences)
                    {
                        summarizedText.Add(sentence.Text);
                        //Console.WriteLine($"  Sentence: {sentence.Text}");
                        //Console.WriteLine();
                    }
                }
            }
        }

    }
        public async static Task<IEnumerable> GenerateSummarizedText(string BodyText, string azureCredentials, string azureCognitiveServiceEndPoint)
        {

            //string keySecretva = keySecret;
            //string endpointval = endpointtest;
            //var client = new TextAnalyticsClient(endpoint, credentials);
            credentials = new AzureKeyCredential(azureCredentials);
            endpoint = new Uri(azureCognitiveServiceEndPoint);
            var client = new TextAnalyticsClient(endpoint, credentials);

            await TextSummarizationExample(client, BodyText);
            return summarizedText;
        }
}
}
