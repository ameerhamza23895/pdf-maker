import React, { createContext, useContext, useState } from "react";

export type ImageFilter = "none" | "mono" | "warm" | "cool";

export type EditableImage = {
  id: string;
  uri: string;
  rotation: number;
  filter: ImageFilter;
  crop: {
    top: number;
    left: number;
    right: number;
    bottom: number;
  };
  processedUri?: string;
};

type EditImagesContextType = {
  images: EditableImage[];
  currentIndex: number;
  addImages: (uris: string[]) => void;
  setCurrentIndex: (index: number) => void;
  updateImage: (id: string, patch: Partial<Omit<EditableImage, "id">>) => void;
  removeImage: (id: string) => void;
  clearImages: () => void;
};

const defaultContext: EditImagesContextType = {
  images: [],
  currentIndex: 0,
  addImages: () => undefined,
  setCurrentIndex: () => undefined,
  updateImage: () => undefined,
  removeImage: () => undefined,
  clearImages: () => undefined,
};

const EditImagesContext = createContext<EditImagesContextType>(defaultContext);

export function EditImagesProvider({
  children,
}: {
  children: React.ReactNode;
}) {
  const [images, setImages] = useState<EditableImage[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);

  const addImages = (uris: string[]) => {
    const newImages = uris.map((uri) => ({
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      uri,
      rotation: 0,
      filter: "none" as ImageFilter,
      crop: { top: 0, left: 0, right: 0, bottom: 0 },
    }));

    setImages((prev) => {
      const next = [...prev, ...newImages];
      return next;
    });

    if (images.length === 0 && newImages.length > 0) {
      setCurrentIndex(0);
    }
  };

  const updateImage = (
    id: string,
    patch: Partial<Omit<EditableImage, "id">>,
  ) => {
    setImages((prev) =>
      prev.map((image) => (image.id === id ? { ...image, ...patch } : image)),
    );
  };

  const removeImage = (id: string) => {
    setImages((prev) => {
      const next = prev.filter((image) => image.id !== id);
      setCurrentIndex((current) => {
        if (next.length === 0) return 0;
        const removedIndex = prev.findIndex((image) => image.id === id);
        if (removedIndex === -1) return current;
        if (removedIndex < current) return current - 1;
        return Math.min(current, next.length - 1);
      });
      return next;
    });
  };

  const clearImages = () => {
    setImages([]);
    setCurrentIndex(0);
  };

  const value = {
    images,
    currentIndex,
    addImages,
    setCurrentIndex,
    updateImage,
    removeImage,
    clearImages,
  };

  return (
    <EditImagesContext.Provider value={value}>
      {children}
    </EditImagesContext.Provider>
  );
}

export function useEditImages() {
  return useContext(EditImagesContext);
}
