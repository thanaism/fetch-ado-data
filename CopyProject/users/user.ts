export type User = {
  displayName: string;
  url: string;
  _links: {
    avatar: {
      href: string;
    };
  };
  id: string;
  uniqueName: string;
  imageUrl: string;
  descriptor: string;
};
