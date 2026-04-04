import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import type { BarcodeScanningResult } from "expo-camera";
import { CameraType, CameraView, useCameraPermissions } from "expo-camera";
import * as ImagePicker from "expo-image-picker";
import { useRouter } from "expo-router";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  ActivityIndicator,
  Alert,
  Button,
  StyleSheet,
  Text,
  TouchableOpacity,
  View,
} from "react-native";
import { useSafeAreaInsets } from "react-native-safe-area-context";

import { QRResultModal } from "@/src/components/QRResultModal";
import { useEditImages } from "@/src/context/edit-images-context";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

const { colors, spacing, radius } = electricCuratorTheme;

/** Single-shot AF locks focus; "off" lets the device adjust continuously (better for QR). */
type AutofocusMode = "on" | "off";

export default function ScanPage() {
  const [facing, setFacing] = useState<CameraType>("back");
  const [flash, setFlash] = useState<"off" | "on">("off");
  const [autofocusMode, setAutofocusMode] = useState<AutofocusMode>("off");
  const [scanEnabled, setScanEnabled] = useState(true);
  const router = useRouter();
  const { addImages } = useEditImages();
  const [permission, requestPermission] = useCameraPermissions();
  const [scannedData, setScannedData] = useState<string | null>(null);
  const [cameraReady, setCameraReady] = useState(false);
  const [capturing, setCapturing] = useState(false);
  const cameraRef = useRef<InstanceType<typeof CameraView>>(null);
  const lastQrAtRef = useRef(0);
  const insets = useSafeAreaInsets();

  useEffect(() => {
    setCameraReady(false);
  }, [facing]);

  const handleBarCodeScanned = useCallback(
    (result: BarcodeScanningResult) => {
      if (!scanEnabled || scannedData) {
        return;
      }
      const now = Date.now();
      if (now - lastQrAtRef.current < 1200) {
        return;
      }
      lastQrAtRef.current = now;
      setScannedData(result.data);
    },
    [scanEnabled, scannedData],
  );

  if (!permission) {
    return (
      <View style={styles.container}>
        <ActivityIndicator size="large" color={colors.primary} />
      </View>
    );
  }

  if (!permission.granted) {
    return (
      <View style={styles.container}>
        <Text style={styles.message}>
          We need your permission to show the camera
        </Text>
        <Button onPress={requestPermission} title="Grant permission" />
      </View>
    );
  }

  function toggleCameraFacing() {
    setFacing((current) => (current === "back" ? "front" : "back"));
  }

  function toggleFlash() {
    setFlash((current) => (current === "off" ? "on" : "off"));
  }

  function toggleScanEnabled() {
    setScanEnabled((current) => !current);
  }

  function toggleAutofocusMode() {
    setAutofocusMode((current) => (current === "off" ? "on" : "off"));
  }

  const takePicture = async () => {
    if (!cameraRef.current || !cameraReady || capturing) {
      return;
    }
    try {
      setCapturing(true);
      const photo = await cameraRef.current.takePictureAsync({
        quality: 0.92,
        skipProcessing: false,
      });
      if (photo?.uri) {
        addImages([photo.uri]);
        router.push("/edit-images");
      }
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Could not capture photo.";
      Alert.alert("Capture failed", message);
    } finally {
      setCapturing(false);
    }
  };

  const pickImage = async () => {
    const { status } = await ImagePicker.requestMediaLibraryPermissionsAsync();

    if (status !== "granted") {
      Alert.alert(
        "Permission Denied",
        "We need access to your photos to make this work!",
      );
      return;
    }

    const result = await ImagePicker.launchImageLibraryAsync({
      mediaTypes: ImagePicker.MediaTypeOptions.Images,
      allowsMultipleSelection: true,
      allowsEditing: false,
      quality: 1,
    });

    if (!result.canceled && result.assets.length > 0) {
      const uris = result.assets.map((asset) => asset.uri);
      addImages(uris);
      router.push("/edit-images");
    }
  };

  return (
    <View style={styles.container}>
      <CameraView
        ref={cameraRef}
        style={styles.camera}
        facing={facing}
        mode="picture"
        enableTorch={flash === "on"}
        autofocus={autofocusMode}
        barcodeScannerSettings={{ barcodeTypes: ["qr"] }}
        onCameraReady={() => setCameraReady(true)}
        onMountError={(event) => {
          setCameraReady(false);
          Alert.alert(
            "Camera error",
            event.message || "Could not start the camera preview.",
          );
        }}
        onBarcodeScanned={
          scanEnabled && !scannedData ? handleBarCodeScanned : undefined
        }
      />

      <QRResultModal
        isVisible={!!scannedData}
        data={scannedData}
        onClose={() => setScannedData(null)}
      />

      <View
        style={[
          styles.topControls,
          { paddingTop: Math.max(insets.top, spacing.sm) },
        ]}
      >
        <TouchableOpacity
          style={[styles.iconButton, scanEnabled && styles.iconButtonActive]}
          onPress={toggleScanEnabled}
        >
          <MaterialIcons
            name="qr-code-scanner"
            size={20}
            color={scanEnabled ? colors.primaryContainer : colors.onSurface}
          />
        </TouchableOpacity>

        <TouchableOpacity style={styles.iconButton} onPress={toggleFlash}>
          <MaterialIcons
            name={flash === "on" ? "flash-on" : "flash-off"}
            size={20}
            color={flash === "on" ? colors.primaryContainer : colors.onSurface}
          />
        </TouchableOpacity>

        <TouchableOpacity
          style={[
            styles.iconButton,
            autofocusMode === "on" && styles.iconButtonActive,
          ]}
          onPress={toggleAutofocusMode}
        >
          <MaterialIcons
            name="center-focus-strong"
            size={20}
            color={
              autofocusMode === "on"
                ? colors.primaryContainer
                : colors.onSurface
            }
          />
        </TouchableOpacity>

        <TouchableOpacity
          style={styles.iconButton}
          onPress={toggleCameraFacing}
        >
          <MaterialIcons
            name="flip-camera-ios"
            size={20}
            color={colors.onSurface}
          />
        </TouchableOpacity>
      </View>

      <View style={styles.focusOverlay} pointerEvents="none">
        <View style={styles.focusBox}>
          <View style={[styles.focusCorner, styles.topLeft]} />
          <View style={[styles.focusCorner, styles.topRight]} />
          <View style={[styles.focusCorner, styles.bottomLeft]} />
          <View style={[styles.focusCorner, styles.bottomRight]} />
        </View>
      </View>

      <View
        style={[
          styles.bottomRow,
          { paddingBottom: Math.max(insets.bottom, spacing.sm) },
        ]}
      >
        <TouchableOpacity style={styles.galleryButton} onPress={pickImage}>
          <MaterialIcons
            name="photo-library"
            size={24}
            color={colors.primary}
          />
        </TouchableOpacity>

        <View style={styles.captureWrapper}>
          <TouchableOpacity
            style={[
              styles.captureButton,
              (!cameraReady || capturing) && styles.captureButtonDisabled,
            ]}
            disabled={!cameraReady || capturing}
            onPress={takePicture}
            accessibilityRole="button"
            accessibilityLabel="Capture photo"
          />
          <Text style={styles.captureLabel}>
            {capturing ? "Saving…" : "Capture"}
          </Text>
        </View>

        <View style={styles.placeholder} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: colors.surface,
  },
  message: {
    textAlign: "center",
    paddingBottom: 10,
  },
  camera: {
    flex: 1,
  },
  topControls: {
    position: "absolute",
    top: 0,
    left: spacing.sm,
    right: spacing.sm,
    flexDirection: "row",
    justifyContent: "space-between",
    zIndex: 10,
  },
  iconButton: {
    width: 44,
    height: 44,
    borderRadius: radius.pill,
    backgroundColor: colors.surfaceContainerLowest,
    alignItems: "center",
    justifyContent: "center",
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 6 },
    shadowOpacity: 0.15,
    shadowRadius: 12,
    elevation: 5,
  },
  iconButtonActive: {
    backgroundColor: colors.primary,
  },
  focusOverlay: {
    ...StyleSheet.absoluteFillObject,
    alignItems: "center",
    justifyContent: "center",
    pointerEvents: "none",
  },
  focusBox: {
    width: 240,
    height: 320,
    borderRadius: 24,
    alignItems: "center",
    justifyContent: "center",
  },
  focusCorner: {
    position: "absolute",
    width: 24,
    height: 24,
    borderColor: colors.surfaceContainerLowest,
  },
  topLeft: {
    top: 0,
    left: 0,
    borderTopWidth: 2,
    borderLeftWidth: 2,
  },
  topRight: {
    top: 0,
    right: 0,
    borderTopWidth: 2,
    borderRightWidth: 2,
  },
  bottomLeft: {
    bottom: 0,
    left: 0,
    borderBottomWidth: 2,
    borderLeftWidth: 2,
  },
  bottomRight: {
    bottom: 0,
    right: 0,
    borderBottomWidth: 2,
    borderRightWidth: 2,
  },
  bottomRow: {
    position: "absolute",
    left: 0,
    right: 0,
    bottom: spacing.sm,
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    paddingHorizontal: spacing.md,
  },
  galleryButton: {
    width: 48,
    height: 48,
    borderRadius: radius.md,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: colors.surfaceContainerLowest,
  },
  captureWrapper: {
    alignItems: "center",
  },
  captureButton: {
    width: 72,
    height: 72,
    borderRadius: 72,
    borderWidth: 4,
    borderColor: colors.surfaceContainerLowest,
    backgroundColor: colors.primary,
  },
  captureButtonDisabled: {
    opacity: 0.45,
  },
  captureLabel: {
    marginTop: 8,
    color: colors.onSurface,
    fontSize: 12,
    fontWeight: "600",
  },
  placeholder: {
    width: 48,
    height: 48,
  },
});
