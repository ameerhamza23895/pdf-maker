import React, { createContext, useContext, useState } from "react";

import type { ImageFilterId } from "@/src/constants/imageFilterPresets";

export type EditableImage = {
  id: string;
  uri: string;
  rotation: number;
  /** Per-image preset; each thumbnail keeps its own value. */
  filter: ImageFilterId;
  crop: {
    top: number;
    left: number;
    right: number;
    bottom: number;
  };
  processedUri?: string;
  key?: string;
};

type EditImagesContextType = {
  images: EditableImage[];
  batchImages: EditableImage[];
  currentIndex: number;
  addImages: (uris: string[]) => void;
  setCurrentIndex: (index: number) => void;
  updateImage: (id: string, patch: Partial<Omit<EditableImage, "id">>) => void;
  removeImage: (id: string) => void;
  clearImages: () => void;
  setImages: (images: EditableImage[]) => void;
  commitBatchImages: (batch: EditableImage[]) => void;
  clearBatchImages: () => void;
};

const EditImagesContext = createContext<EditImagesContextType | undefined>(
  undefined,
);

export function EditImagesProvider({
  children,
}: {
  children: React.ReactNode;
}) {
  const [images, setImagesState] = useState<EditableImage[]>([]);
  const [batchImages, setBatchImagesState] = useState<EditableImage[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);

  const addImages = (uris: string[]) => {
    const newBatch = uris.map((uri, index) => ({
      id: `${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
      uri,
      rotation: 0,
      filter: "none" as ImageFilterId,
      crop: { top: 0, left: 0, right: 0, bottom: 0 },
      key: `${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    }));
    setBatchImagesState(newBatch);
    setCurrentIndex(0);
  };

  const updateImage = (
    id: string,
    patch: Partial<Omit<EditableImage, "id">>,
  ) => {
    setBatchImagesState((prev) =>
      prev.map((image) => (image.id === id ? { ...image, ...patch } : image)),
    );
  };

  const setImages = (newImages: EditableImage[]) => {
    setImagesState(newImages);
  };

  const commitBatchImages = (batch: EditableImage[]) => {
    setImagesState((prev) => [...prev, ...batch]);
    setBatchImagesState([]);
    setCurrentIndex(0);
  };

  const clearBatchImages = () => {
    setBatchImagesState([]);
    setCurrentIndex(0);
  };

  const removeImage = (id: string) => {
    setBatchImagesState((prev) => {
      const next = prev.filter((image) => image.id !== id);
      setCurrentIndex((current) => {
        if (next.length === 0) return 0;
        const removedIndex = prev.findIndex((image) => image.id === id);
        if (removedIndex < current) return current - 1;
        return Math.min(current, next.length - 1);
      });
      return next;
    });
  };

  const clearImages = () => {
    setImagesState([]);
    setBatchImagesState([]);
    setCurrentIndex(0);
  };

  return (
    <EditImagesContext.Provider
      value={{
        images,
        batchImages,
        currentIndex,
        addImages,
        setCurrentIndex,
        updateImage,
        removeImage,
        clearImages,
        setImages,
        commitBatchImages,
        clearBatchImages,
      }}
    >
      {children}
    </EditImagesContext.Provider>
  );
}

export function useEditImages() {
  const context = useContext(EditImagesContext);
  if (!context)
    throw new Error("useEditImages must be used within EditImagesProvider");
  return context;
}
