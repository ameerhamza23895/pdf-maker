import { useEffect } from "react";

type Props = {
  visible?: boolean;
  onVisibilityChange?: (visible: boolean) => void;
};

export function AdBanner({
  visible = true,
  onVisibilityChange,
}: Props) {
  useEffect(() => {
    onVisibilityChange?.(false);
  }, [onVisibilityChange, visible]);

  return null;
}
