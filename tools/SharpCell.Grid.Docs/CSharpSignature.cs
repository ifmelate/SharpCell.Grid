using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;

namespace SharpCell.Grid.Docs;

/// <summary>How a type or member is declared, written as C#.</summary>
internal static class CSharpSignature
{
    // NullabilityInfoContext caches into a plain Dictionary and is not thread-safe; tests run in parallel.
    private static NullabilityInfoContext Nullability => new();

    // By full name: the types are read as metadata, so they are not the runtime's own.
    private static readonly Dictionary<string, string> Keywords = new(StringComparer.Ordinal)
    {
        ["System.Void"] = "void", ["System.Object"] = "object", ["System.String"] = "string", ["System.Boolean"] = "bool",
        ["System.Char"] = "char", ["System.Byte"] = "byte", ["System.Int16"] = "short", ["System.Int32"] = "int",
        ["System.Int64"] = "long", ["System.UInt32"] = "uint", ["System.UInt64"] = "ulong", ["System.Single"] = "float",
        ["System.Double"] = "double", ["System.Decimal"] = "decimal",
    };

    private static readonly Dictionary<string, string> Operators = new(StringComparer.Ordinal)
    {
        ["op_Equality"] = "==", ["op_Inequality"] = "!=", ["op_LessThan"] = "<", ["op_GreaterThan"] = ">",
        ["op_LessThanOrEqual"] = "<=", ["op_GreaterThanOrEqual"] = ">=", ["op_Addition"] = "+", ["op_Subtraction"] = "-",
    };

    public static string Declaration(Type type)
    {
        var sb = new StringBuilder("public ");
        if (type.IsEnum)
            return sb.Append("enum ").Append(type.Name).ToString();
        if (type.IsValueType)
        {
            if (ApiReference.Has(type, "System.Runtime.CompilerServices.IsReadOnlyAttribute"))
                sb.Append("readonly ");
            return sb.Append("struct ").Append(type.Name).ToString();
        }

        if (type.IsAbstract && type.IsSealed)
            sb.Append("static ");
        else if (type.IsSealed)
            sb.Append("sealed ");
        else if (type.IsAbstract)
            sb.Append("abstract ");
        sb.Append("class ").Append(type.Name);
        if (type.BaseType is { } baseType && baseType.FullName != "System.Object")
            sb.Append(" : ").Append(TypeName(baseType));
        return sb.ToString();
    }

    public static string Of(MemberInfo member) => member switch
    {
        ConstructorInfo constructor => $"public {constructor.DeclaringType!.Name}({Parameters(constructor.GetParameters(), false)})",
        MethodInfo method => Method(method),
        PropertyInfo property => Property(property),
        FieldInfo { DeclaringType.IsEnum: true } field =>
            $"{field.Name} = {Convert.ToInt64(field.GetRawConstantValue(), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)}",
        FieldInfo field => $"public {(field.IsLiteral ? "const" : field.IsStatic ? "static readonly" : "readonly")} {TypeName(field.FieldType, Nullability.Create(field))} {field.Name}",
        _ => throw new NotSupportedException($"No signature for {member.MemberType} {member.Name}."),
    };

    /// <summary>The member as a heading: name and parameter types, enough to tell overloads apart.</summary>
    public static string Heading(MemberInfo member) => member switch
    {
        ConstructorInfo constructor => $"{constructor.DeclaringType!.Name}({ParameterTypes(constructor.GetParameters())})",
        MethodInfo { Name: "op_Implicit" or "op_Explicit" } conversion =>
            $"{(conversion.Name == "op_Implicit" ? "implicit" : "explicit")} operator {TypeName(conversion.ReturnType)}({ParameterTypes(conversion.GetParameters())})",
        MethodInfo method when Operators.TryGetValue(method.Name, out var symbol) => $"operator {symbol}({ParameterTypes(method.GetParameters())})",
        MethodInfo method => $"{method.Name}({ParameterTypes(method.GetParameters())})",
        PropertyInfo property when property.GetIndexParameters().Length > 0 => $"this[{ParameterTypes(property.GetIndexParameters())}]",
        _ => member.Name,
    };

    /// <summary>A stable HTML id for the member, unique among overloads and operators.</summary>
    public static string Anchor(MemberInfo member)
    {
        var parameters = member switch
        {
            MethodBase method => method.GetParameters(),
            PropertyInfo property => property.GetIndexParameters(),
            _ => [],
        };
        var name = member is ConstructorInfo ? "ctor" : member.Name;
        return Slug(string.Join(" ", parameters.Select(p => TypeName(p.ParameterType)).Prepend(name)));
    }

    public static string TypeName(Type type, NullabilityInfo? nullability = null)
    {
        if (type.IsByRef)
            return TypeName(type.GetElementType()!, nullability);
        if (type.IsGenericType && !type.IsGenericTypeDefinition && type.GetGenericTypeDefinition().FullName == "System.Nullable`1")
            return TypeName(type.GetGenericArguments()[0]) + "?";

        string name;
        if (type.IsArray)
            name = TypeName(type.GetElementType()!, nullability?.ElementType) + "[" + new string(',', type.GetArrayRank() - 1) + "]";
        else if (type.FullName is { } fullName && Keywords.TryGetValue(fullName, out var keyword))
            name = keyword;
        else if (type.IsGenericType)
        {
            var arguments = type.GetGenericArguments();
            name = type.Name[..type.Name.IndexOf('`', StringComparison.Ordinal)] + "<"
                + string.Join(", ", arguments.Select((a, i) => TypeName(a, nullability?.GenericTypeArguments.ElementAtOrDefault(i)))) + ">";
        }
        else
            name = type.Name;

        return !type.IsValueType && nullability?.ReadState == NullabilityState.Nullable ? name + "?" : name;
    }

    private static string Method(MethodInfo method)
    {
        var modifiers = "public " + (method.IsStatic ? "static " : "")
            + (IsOverride(method) ? "override " : "");
        var parameters = Parameters(method.GetParameters(), ApiReference.Has(method, "System.Runtime.CompilerServices.ExtensionAttribute"));
        if (method.Name is "op_Implicit" or "op_Explicit")
            return $"{modifiers}{(method.Name == "op_Implicit" ? "implicit" : "explicit")} operator {TypeName(method.ReturnType)}({parameters})";
        var returns = TypeName(method.ReturnType, Nullability.Create(method.ReturnParameter));
        return Operators.TryGetValue(method.Name, out var symbol)
            ? $"{modifiers}{returns} operator {symbol}({parameters})"
            : $"{modifiers}{returns} {method.Name}({parameters})";
    }

    // A virtual method that takes no new slot overrides one of a base class.
    private static bool IsOverride(MethodInfo method) =>
        !method.IsStatic && method.IsVirtual && (method.Attributes & MethodAttributes.NewSlot) == 0;

    private static string Property(PropertyInfo property)
    {
        var accessor = property.GetMethod ?? property.SetMethod!;
        var sb = new StringBuilder("public ");
        if (accessor.IsStatic)
            sb.Append("static ");
        sb.Append(TypeName(property.PropertyType, Nullability.Create(property))).Append(' ');
        var index = property.GetIndexParameters();
        sb.Append(index.Length > 0 ? $"this[{Parameters(index, false)}]" : property.Name);
        sb.Append(" { ");
        if (property.GetMethod is { IsPublic: true })
            sb.Append("get; ");
        if (property.SetMethod is { IsPublic: true })
            sb.Append(property.SetMethod.ReturnParameter.GetRequiredCustomModifiers().Any(m => m.FullName == "System.Runtime.CompilerServices.IsExternalInit") ? "init; " : "set; ");
        return sb.Append('}').ToString();
    }

    private static string Parameters(ParameterInfo[] parameters, bool extension) =>
        string.Join(", ", parameters.Select((p, i) => (extension && i == 0 ? "this " : "") + Parameter(p)));

    private static string Parameter(ParameterInfo parameter)
    {
        var prefix = parameter.IsOut ? "out " : parameter.ParameterType.IsByRef ? "ref " : "";
        var text = $"{prefix}{TypeName(parameter.ParameterType, Nullability.Create(parameter))} {parameter.Name}";
        return parameter.HasDefaultValue ? text + " = " + DefaultValue(parameter) : text;
    }

    // Metadata gives an enum's default as its number; its name is the field with that value.
    private static string DefaultValue(ParameterInfo parameter) => parameter.RawDefaultValue switch
    {
        null => parameter.ParameterType.IsValueType ? "default" : "null",
        bool b => b ? "true" : "false",
        string s => "\"" + s.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"",
        var value when parameter.ParameterType.IsEnum => parameter.ParameterType.Name + "."
            + parameter.ParameterType.GetFields(BindingFlags.Public | BindingFlags.Static).First(f => Equals(f.GetRawConstantValue(), value)).Name,
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        var other => other.ToString() ?? "default",
    };

    private static string ParameterTypes(ParameterInfo[] parameters) => string.Join(", ", parameters.Select(p => TypeName(p.ParameterType)));

    private static string Slug(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text.ToLowerInvariant())
        {
            if (c is >= 'a' and <= 'z' or >= '0' and <= '9')
                sb.Append(c);
            else if (sb.Length > 0 && sb[^1] != '-')
                sb.Append('-');
        }

        return sb.ToString().Trim('-');
    }
}
