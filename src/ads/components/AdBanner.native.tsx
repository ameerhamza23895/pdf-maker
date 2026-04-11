import { useEffect, useMemo, useState } from "react";
import { StyleSheet, View } from "react-native";

import { getBannerUnitIdForRuntime } from "../config/resolveAdConfig";
import { getGoogleMobileAds } from "../googleMobileAds";
import { electricCuratorTheme, withAlpha } from "@/src/theme/electric-curator";

const { colors, radius, spacing } = electricCuratorTheme;
const googleMobileAds = getGoogleMobileAds();

type Props = {
  visible?: boolean;
  onVisibilityChange?: (visible: boolean) => void;
};

export function AdBanner({
  visible = true,
  onVisibilityChange,
}: Props) {
  const [isLoaded, setIsLoaded] = useState(false);
  const unitId = useMemo(() => getBannerUnitIdForRuntime(), []);

  useEffect(() => {
    onVisibilityChange?.(visible && isLoaded);
  }, [isLoaded, onVisibilityChange, visible]);

  useEffect(() => {
    if (!visible) {
      setIsLoaded(false);
    }
  }, [visible]);

  if (!visible || !googleMobileAds) {
    return null;
  }

  const { BannerAd, BannerAdSize } = googleMobileAds;

  return (
    <View style={[styles.wrap, !isLoaded && styles.hidden]}>
      <BannerAd
        unitId={unitId}
        size={BannerAdSize.ANCHORED_ADAPTIVE_BANNER}
        requestOptions={{ requestNonPersonalizedAdsOnly: true }}
        onAdLoaded={() => setIsLoaded(true)}
        onAdFailedToLoad={() => setIsLoaded(false)}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    alignItems: "center",
    gap: spacing.xs,
    paddingVertical: spacing.sm,
    borderRadius: radius.md,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.outlineVariant, 0.5),
    backgroundColor: colors.surfaceContainerLowest,
  },
  hidden: {
    height: 0,
    opacity: 0,
    overflow: "hidden",
    paddingVertical: 0,
    borderWidth: 0,
  },
});
