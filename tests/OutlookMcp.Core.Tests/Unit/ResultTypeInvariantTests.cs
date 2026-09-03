using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using OutlookMcp.Core.Models;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Validates invariants that every result type has to hold, to prevent Rule 1 violations
/// (Success=true with ErrorMessage set) and null collections reaching a caller.
///
/// <para>
/// These are driven by reflection over the whole model namespace rather than written out once per
/// type. The hand-written version asserted the same two things twenty-four times, every one of them
/// against a PowerPoint type inherited from the fork; when those types were deleted the coverage
/// went with them, and no Outlook result type had ever been checked at all. Reflection covers every
/// type that exists now and every one added later, which is the only version of this test that
/// cannot quietly stop testing anything.
/// </para>
///
/// <para>
/// This is the documented Rule 30 exception: pure reflection over plain objects, no COM, no Outlook.
/// </para>
/// </summary>
public class ResultTypeInvariantTests
{
    /// <summary>
    /// Every constructible public model type. Abstract types and those needing constructor arguments
    /// are skipped - they cannot be default-constructed, so the invariants do not apply to them.
    /// </summary>
    public static TheoryData<Type> ResultTypes
    {
        get
        {
            var data = new TheoryData<Type>();
            foreach (Type type in TypesUnderTest())
            {
                data.Add(type);
            }

            return data;
        }
    }

    private static IEnumerable<Type> TypesUnderTest() =>
        typeof(OperationResult).Assembly
            .GetTypes()
            .Where(t => t.IsClass
                && t.IsPublic
                && !t.IsAbstract
                && t.Namespace == typeof(OperationResult).Namespace
                && t.GetConstructor(Type.EmptyTypes) != null)
            .OrderBy(t => t.Name, StringComparer.Ordinal);

    /// <summary>
    /// A result must never default to claiming success. Rule 1's invariant is that Success=true
    /// implies no error; a type that starts out true means any path which forgets to set it reports
    /// a success that never happened - the failure mode this project keeps finding.
    /// </summary>
    [Theory]
    [MemberData(nameof(ResultTypes))]
    public void ResultType_DoesNotDefaultToSuccess(Type type)
    {
        PropertyInfo? success = type.GetProperty("Success", BindingFlags.Public | BindingFlags.Instance);

        if (success is null || success.PropertyType != typeof(bool))
        {
            return;
        }

        object instance = Activator.CreateInstance(type)!;

        Assert.False(
            (bool)success.GetValue(instance)!,
            $"{type.Name}.Success defaults to true, so a code path that never sets it reports success.");
    }

    /// <summary>
    /// A collection property that is not declared nullable must be an empty list, never null. A
    /// caller that has to null-check every collection will eventually not.
    ///
    /// <para>
    /// Properties declared <c>List&lt;T&gt;?</c> are exempt, and the distinction is deliberate rather
    /// than a loophole: those are tri-state on the wire. <c>MailSummaryInfo.AccessDenied</c> is null
    /// when nothing was blocked and is then omitted entirely, because an empty list would serialise
    /// <c>"accessDenied": []</c> onto every message in every listing. Declaring the property nullable
    /// is how that intent is stated, so this test reads it rather than overriding it.
    /// </para>
    /// </summary>
    [Theory]
    [MemberData(nameof(ResultTypes))]
    public void ResultType_NonNullableCollectionsAreEmptyNotNull(Type type)
    {
        object instance = Activator.CreateInstance(type)!;
        var nullability = new NullabilityInfoContext();

        PropertyInfo[] collections = [.. type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead
                && p.PropertyType != typeof(string)
                && typeof(IEnumerable).IsAssignableFrom(p.PropertyType)
                && nullability.Create(p).ReadState != NullabilityState.Nullable)];

        foreach (PropertyInfo property in collections)
        {
            object? value = property.GetValue(instance);

            Assert.True(
                value is not null,
                $"{type.Name}.{property.Name} is a non-nullable collection but defaults to null. "
                + "Initialise it to an empty collection, or declare it nullable if null is meaningful.");

            Assert.False(
                ((IEnumerable)value!).GetEnumerator().MoveNext(),
                $"{type.Name}.{property.Name} is not empty by default.");
        }
    }

    /// <summary>
    /// The namespace must actually contain result types, and they must actually have collections and
    /// Success flags to check. Without this, a reflection filter that stopped matching would leave a
    /// suite that passes with zero assertions - which is exactly how the per-type version failed.
    /// </summary>
    [Fact]
    public void TypesUnderTest_CoverTheModelSurface()
    {
        List<Type> types = [.. TypesUnderTest()];

        Assert.True(
            types.Count > 20,
            $"Only {types.Count} model types were discovered; the reflection filter has stopped matching.");

        Assert.True(
            types.Count(t => t.GetProperty("Success") is not null) > 20,
            "Almost no discovered type has a Success flag; the filter is matching the wrong things.");

        Assert.True(
            types.Any(t => t.GetProperties()
                .Any(p => p.PropertyType != typeof(string)
                    && typeof(IEnumerable).IsAssignableFrom(p.PropertyType)
                    && new NullabilityInfoContext().Create(p).ReadState != NullabilityState.Nullable)),
            "No discovered type has a non-nullable collection property, so that invariant asserts nothing.");
    }

    [Fact]
    public void OperationResult_DefaultState_SuccessIsFalse()
    {
        var result = new OperationResult();
        Assert.False(result.Success);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public void OperationResult_SuccessTrue_ErrorMessageMustBeNull()
    {
        var result = new OperationResult { Success = true };
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
    }
}
