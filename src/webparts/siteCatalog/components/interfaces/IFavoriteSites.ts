export interface IFavoriteSites {
    Items: Array<FavoriteSite>;
}

export interface FavoriteSite {
    ItemReference: {
        IndexId: number;
        Type: string;
    };
    BannerColor: string;
    Url: string;
    Title: string;
    Type: string;
}