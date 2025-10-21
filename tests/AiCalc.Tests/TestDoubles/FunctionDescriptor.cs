using System;
using System.Collections.Generic;
using AiCalc.Models;

namespace AiCalc.Services
{
    public enum FunctionCategory
    {
        Math,
        Text,
        DateTime,
        File,
        Directory,
        Table,
        Image,
        Video,
        Pdf,
        Data,
        AI,
        Contrib
    }

    public class FunctionParameter
    {
        public FunctionParameter(string name, string description, CellObjectType expectedType, bool isOptional = false, params CellObjectType[] additionalAcceptableTypes)
        {
            Name = name;
            Description = description;
            ExpectedType = expectedType;
            IsOptional = isOptional;
            var list = new List<CellObjectType> { expectedType };
            if (additionalAcceptableTypes != null)
            {
                list.AddRange(additionalAcceptableTypes);
            }
            AcceptableTypes = list;
        }

        public string Name { get; }
        public string Description { get; }
        public CellObjectType ExpectedType { get; }
        public bool IsOptional { get; }
        public IReadOnlyList<CellObjectType> AcceptableTypes { get; }

        public bool CanAccept(CellObjectType type) => AcceptableTypes.Contains(type);
    }

    public class FunctionDescriptor
    {
        public FunctionDescriptor(string name, string description, Func<object, System.Threading.Tasks.Task<object>> handler, FunctionCategory category = FunctionCategory.Math, params FunctionParameter[] parameters)
        {
            Name = name;
            Description = description;
            Category = category;
            Parameters = parameters;
        }

        public string Name { get; }
        public string Description { get; }
        public FunctionCategory Category { get; }
        public IReadOnlyList<FunctionParameter> Parameters { get; }
    }
}
