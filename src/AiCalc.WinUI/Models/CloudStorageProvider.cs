using System.Runtime.Serialization;

namespace AiCalc.Models;

public enum CloudStorageProvider
{
    [EnumMember(Value = "AzureBlob")]
    AzureBlob,
    [EnumMember(Value = "AwsS3")]
    AwsS3,
    [EnumMember(Value = "GoogleCloudStorage")]
    GoogleCloudStorage
}
