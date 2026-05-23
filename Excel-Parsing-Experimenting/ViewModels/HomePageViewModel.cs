namespace Excel_Parsing_Experimenting.ViewModels;

public sealed class HomePageViewModel
{
    public IReadOnlyList<ParserCardViewModel> ParserCards { get; init; } = [];
}

public sealed class ParserCardViewModel
{
    public required string DisplayName { get; init; }

    public required string Summary { get; init; }

    public required string LibraryUrl { get; init; }

    public required string UploadAction { get; init; }

    public required IReadOnlyList<string> SupportedExtensions { get; init; }

    public required IReadOnlyList<string> Pros { get; init; }

    public required IReadOnlyList<string> Cons { get; init; }

    public string InputId => $"upload-{DisplayName.ToLowerInvariant().Replace(' ', '-')}";
}
