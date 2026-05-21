import React, {useState, useEffect, useCallback, useRef} from 'react';
import {
  SafeAreaView,
  StyleSheet,
  Text,
  View,
  TouchableOpacity,
  Alert,
  Platform,
  PermissionsAndroid,
  Linking,
  AppState,
  AppStateStatus,
} from 'react-native';
import Geolocation from 'react-native-geolocation-service';

// GASのWebアプリURL（デプロイ後に設定）
const GAS_WEB_APP_URL = 'YOUR_GAS_WEB_APP_URL_HERE';

// 位置情報送信間隔（ミリ秒）
const LOCATION_INTERVAL = 30000; // 30秒

type LocationData = {
  latitude: number;
  longitude: number;
  timestamp: number;
  accuracy: number | null;
  speed: number | null;
};

const App = () => {
  const [isTracking, setIsTracking] = useState(false);
  const [lastLocation, setLastLocation] = useState<LocationData | null>(null);
  const [statusMessage, setStatusMessage] = useState('待機中');
  const watchIdRef = useRef<number | null>(null);
  const intervalRef = useRef<NodeJS.Timeout | null>(null);
  const appState = useRef(AppState.currentState);

  // Deep Link処理
  const handleDeepLink = useCallback((url: string | null) => {
    if (url && url.startsWith('deliverytracker://start')) {
      console.log('Deep Link received:', url);
      // すでに送信中でなければ開始
      if (!isTracking) {
        startTracking();
      }
    }
  }, [isTracking]);

  // Deep Linkの初期化とリスナー設定
  useEffect(() => {
    // アプリ起動時のDeep Link確認
    const getInitialURL = async () => {
      const initialUrl = await Linking.getInitialURL();
      if (initialUrl) {
        handleDeepLink(initialUrl);
      }
    };
    getInitialURL();

    // Deep Linkリスナー
    const linkingSubscription = Linking.addEventListener('url', (event) => {
      handleDeepLink(event.url);
    });

    // AppState監視（バックグラウンド復帰時のDeep Link対応）
    const appStateSubscription = AppState.addEventListener('change', (nextAppState: AppStateStatus) => {
      if (appState.current.match(/inactive|background/) && nextAppState === 'active') {
        // フォアグラウンドに戻った時
        Linking.getInitialURL().then(handleDeepLink);
      }
      appState.current = nextAppState;
    });

    return () => {
      linkingSubscription.remove();
      appStateSubscription.remove();
    };
  }, [handleDeepLink]);

  // 位置情報権限リクエスト
  const requestLocationPermission = async (): Promise<boolean> => {
    if (Platform.OS === 'android') {
      try {
        const fineLocation = await PermissionsAndroid.request(
          PermissionsAndroid.PERMISSIONS.ACCESS_FINE_LOCATION,
          {
            title: '位置情報の許可',
            message: '配送トラッキングのために位置情報へのアクセスが必要です。',
            buttonNeutral: '後で',
            buttonNegative: 'キャンセル',
            buttonPositive: '許可',
          },
        );

        if (fineLocation === PermissionsAndroid.RESULTS.GRANTED) {
          // バックグラウンド位置情報も要求
          const backgroundLocation = await PermissionsAndroid.request(
            PermissionsAndroid.PERMISSIONS.ACCESS_BACKGROUND_LOCATION,
            {
              title: 'バックグラウンド位置情報の許可',
              message: 'アプリがバックグラウンドでも位置情報を送信できるようにしてください。',
              buttonNeutral: '後で',
              buttonNegative: 'キャンセル',
              buttonPositive: '許可',
            },
          );
          return backgroundLocation === PermissionsAndroid.RESULTS.GRANTED;
        }
        return false;
      } catch (err) {
        console.warn(err);
        return false;
      }
    }
    return true;
  };

  // 位置情報を取得
  const getCurrentLocation = (): Promise<LocationData> => {
    return new Promise((resolve, reject) => {
      Geolocation.getCurrentPosition(
        (position) => {
          resolve({
            latitude: position.coords.latitude,
            longitude: position.coords.longitude,
            timestamp: position.timestamp,
            accuracy: position.coords.accuracy,
            speed: position.coords.speed,
          });
        },
        (error) => {
          reject(error);
        },
        {
          enableHighAccuracy: true,
          timeout: 15000,
          maximumAge: 10000,
        },
      );
    });
  };

  // 位置情報をサーバーに送信
  const sendLocationToServer = async (location: LocationData) => {
    try {
      const response = await fetch(GAS_WEB_APP_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: 'updateLocation',
          data: location,
        }),
      });

      if (response.ok) {
        setStatusMessage(`送信成功: ${new Date().toLocaleTimeString()}`);
      } else {
        setStatusMessage('送信エラー');
      }
    } catch (error) {
      console.error('Location send error:', error);
      setStatusMessage('通信エラー');
    }
  };

  // トラッキング開始
  const startTracking = async () => {
    const hasPermission = await requestLocationPermission();
    if (!hasPermission) {
      Alert.alert('権限エラー', '位置情報の権限が必要です。設定から許可してください。');
      return;
    }

    setIsTracking(true);
    setStatusMessage('トラッキング中...');

    // 初回位置取得
    try {
      const location = await getCurrentLocation();
      setLastLocation(location);
      await sendLocationToServer(location);
    } catch (error) {
      console.error('Initial location error:', error);
    }

    // 定期的な位置情報送信
    intervalRef.current = setInterval(async () => {
      try {
        const location = await getCurrentLocation();
        setLastLocation(location);
        await sendLocationToServer(location);
      } catch (error) {
        console.error('Interval location error:', error);
        setStatusMessage('位置取得エラー');
      }
    }, LOCATION_INTERVAL);
  };

  // トラッキング停止
  const stopTracking = () => {
    if (intervalRef.current) {
      clearInterval(intervalRef.current);
      intervalRef.current = null;
    }
    if (watchIdRef.current !== null) {
      Geolocation.clearWatch(watchIdRef.current);
      watchIdRef.current = null;
    }
    setIsTracking(false);
    setStatusMessage('停止中');
  };

  // クリーンアップ
  useEffect(() => {
    return () => {
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }
      if (watchIdRef.current !== null) {
        Geolocation.clearWatch(watchIdRef.current);
      }
    };
  }, []);

  return (
    <SafeAreaView style={styles.container}>
      <View style={styles.header}>
        <Text style={styles.title}>配送トラッカー</Text>
        <Text style={styles.subtitle}>位置情報送信アプリ</Text>
      </View>

      <View style={styles.statusContainer}>
        <View style={[styles.statusIndicator, isTracking ? styles.statusActive : styles.statusInactive]} />
        <Text style={styles.statusText}>{statusMessage}</Text>
      </View>

      {lastLocation && (
        <View style={styles.locationInfo}>
          <Text style={styles.locationLabel}>最新位置情報:</Text>
          <Text style={styles.locationText}>
            緯度: {lastLocation.latitude.toFixed(6)}
          </Text>
          <Text style={styles.locationText}>
            経度: {lastLocation.longitude.toFixed(6)}
          </Text>
          <Text style={styles.locationText}>
            精度: {lastLocation.accuracy?.toFixed(0) ?? '-'}m
          </Text>
          <Text style={styles.locationText}>
            時刻: {new Date(lastLocation.timestamp).toLocaleTimeString()}
          </Text>
        </View>
      )}

      <View style={styles.buttonContainer}>
        {!isTracking ? (
          <TouchableOpacity style={styles.startButton} onPress={startTracking}>
            <Text style={styles.buttonText}>送信開始</Text>
          </TouchableOpacity>
        ) : (
          <TouchableOpacity style={styles.stopButton} onPress={stopTracking}>
            <Text style={styles.buttonText}>送信停止</Text>
          </TouchableOpacity>
        )}
      </View>

      <View style={styles.footer}>
        <Text style={styles.footerText}>
          Deep Link: deliverytracker://start
        </Text>
      </View>
    </SafeAreaView>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#f5f5f5',
  },
  header: {
    backgroundColor: '#667eea',
    padding: 20,
    alignItems: 'center',
  },
  title: {
    fontSize: 24,
    fontWeight: 'bold',
    color: '#fff',
  },
  subtitle: {
    fontSize: 14,
    color: 'rgba(255,255,255,0.8)',
    marginTop: 4,
  },
  statusContainer: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 20,
    backgroundColor: '#fff',
    margin: 16,
    borderRadius: 12,
    shadowColor: '#000',
    shadowOffset: {width: 0, height: 2},
    shadowOpacity: 0.1,
    shadowRadius: 4,
    elevation: 3,
  },
  statusIndicator: {
    width: 16,
    height: 16,
    borderRadius: 8,
    marginRight: 10,
  },
  statusActive: {
    backgroundColor: '#4CAF50',
  },
  statusInactive: {
    backgroundColor: '#9E9E9E',
  },
  statusText: {
    fontSize: 18,
    color: '#333',
  },
  locationInfo: {
    backgroundColor: '#fff',
    margin: 16,
    marginTop: 0,
    padding: 16,
    borderRadius: 12,
    shadowColor: '#000',
    shadowOffset: {width: 0, height: 2},
    shadowOpacity: 0.1,
    shadowRadius: 4,
    elevation: 3,
  },
  locationLabel: {
    fontSize: 14,
    color: '#666',
    marginBottom: 8,
  },
  locationText: {
    fontSize: 16,
    color: '#333',
    marginBottom: 4,
  },
  buttonContainer: {
    padding: 16,
    marginTop: 'auto',
  },
  startButton: {
    backgroundColor: '#4CAF50',
    padding: 20,
    borderRadius: 12,
    alignItems: 'center',
  },
  stopButton: {
    backgroundColor: '#f44336',
    padding: 20,
    borderRadius: 12,
    alignItems: 'center',
  },
  buttonText: {
    color: '#fff',
    fontSize: 20,
    fontWeight: 'bold',
  },
  footer: {
    padding: 16,
    alignItems: 'center',
  },
  footerText: {
    fontSize: 12,
    color: '#999',
  },
});

export default App;
