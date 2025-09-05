```
import React, { createContext, useContext, useState, useEffect, useRef, useCallback } from 'react';

interface ChartContextValue {
  activeChartId: string | null;
  isChartActive: boolean;
  activateChart: (chartId: string) => Promise<void>;
  deactivateChart: () => void;
  refreshChartSelection: () => Promise<void>;
}

const ChartContext = createContext<ChartContextValue | null>(null);

export const useChart = () => {
  const context = useContext(ChartContext);
  if (!context) {
    throw new Error('useChart must be used within a ChartProvider');
  }
  return context;
};

interface ChartProviderProps {
  children: React.ReactNode;
}

export const ChartProvider: React.FC<ChartProviderProps> = ({ children }) => {
  const [activeChartId, setActiveChartId] = useState<string | null>(null);
  const [isChartActive, setIsChartActive] = useState(false);
  
  // Event handler references to manage cleanup
  const eventHandlers = useRef<Map<string, Excel.EventHandlerResult<any>>>(new Map());
  const isOnlineRef = useRef<boolean>(false);
  const reconnectionTimer = useRef<NodeJS.Timeout | null>(null);

  // Detect if running in Excel Online vs Desktop
  useEffect(() => {
    const detectPlatform = async () => {
      try {
        await Excel.run(async (context) => {
          context.workbook.load('name');
          await context.sync();
          
          // Excel Online specific detection
          isOnlineRef.current = window.location.hostname.includes('office') || 
                               window.location.hostname.includes('sharepoint') ||
                               !!(window as any).OfficeRuntime?.platform === 'web';
        });
      } catch (error) {
        console.error('Platform detection failed:', error);
      }
    };
    
    detectPlatform();
  }, []);

  // Enhanced chart activation handler
  const handleChartActivated = useCallback(async (event: Excel.ChartActivatedEventArgs) => {
    try {
      await Excel.run(async (context) => {
        // Get the activated chart
        const chart = context.workbook.charts.getItem(event.chartId);
        chart.load(['id', 'name']);
        
        await context.sync();
        
        setActiveChartId(chart.id);
        setIsChartActive(true);
        
        console.log(`Chart activated: ${chart.id} (${chart.name})`);
        
        // For Excel Online, implement additional focus management
        if (isOnlineRef.current) {
          // Clear any existing reconnection timer
          if (reconnectionTimer.current) {
            clearTimeout(reconnectionTimer.current);
          }
          
          // Set up periodic check to maintain chart selection
          reconnectionTimer.current = setTimeout(() => {
            refreshChartSelection();
          }, 100); // Quick recheck after activation
        }
      });
    } catch (error) {
      console.error('Chart activation handler failed:', error);
      // Retry activation if it fails
      setTimeout(() => refreshChartSelection(), 500);
    }
  }, []);

  // Enhanced chart deactivation handler
  const handleChartDeactivated = useCallback(async (event: Excel.ChartDeactivatedEventArgs) => {
    try {
      console.log(`Chart deactivated: ${event.chartId}`);
      
      // For Excel Online, add a delay before deactivating to prevent premature deactivation
      const deactivateTimeout = isOnlineRef.current ? 200 : 0;
      
      setTimeout(() => {
        setActiveChartId(null);
        setIsChartActive(false);
        
        // Clear reconnection timer
        if (reconnectionTimer.current) {
          clearTimeout(reconnectionTimer.current);
          reconnectionTimer.current = null;
        }
      }, deactivateTimeout);
      
    } catch (error) {
      console.error('Chart deactivation handler failed:', error);
    }
  }, []);

  // Register event handlers for all charts
  const registerChartEvents = useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        // Clear existing event handlers
        eventHandlers.current.forEach((handler, chartId) => {
          try {
            handler.remove();
          } catch (error) {
            console.warn(`Failed to remove handler for chart ${chartId}:`, error);
          }
        });
        eventHandlers.current.clear();

        // Get all charts in the workbook
        const charts = context.workbook.charts;
        charts.load(['items']);
        
        await context.sync();

        // Register events for each chart
        for (const chart of charts.items) {
          chart.load(['id', 'name']);
          await context.sync();
          
          try {
            // Register onActivated event
            const activatedHandler = chart.onActivated.add(handleChartActivated);
            const deactivatedHandler = chart.onDeactivated.add(handleChartDeactivated);
            
            await context.sync();
            
            // Store handlers for cleanup
            eventHandlers.current.set(`${chart.id}-activated`, activatedHandler);
            eventHandlers.current.set(`${chart.id}-deactivated`, deactivatedHandler);
            
            console.log(`Events registered for chart: ${chart.id} (${chart.name})`);
          } catch (error) {
            console.error(`Failed to register events for chart ${chart.id}:`, error);
          }
        }
      });
    } catch (error) {
      console.error('Failed to register chart events:', error);
    }
  }, [handleChartActivated, handleChartDeactivated]);

  // Manual chart activation
  const activateChart = useCallback(async (chartId: string) => {
    try {
      await Excel.run(async (context) => {
        const chart = context.workbook.charts.getItem(chartId);
        
        // Activate the chart
        chart.activate();
        
        await context.sync();
        
        // Update state immediately for better UX
        setActiveChartId(chartId);
        setIsChartActive(true);
      });
    } catch (error) {
      console.error('Manual chart activation failed:', error);
      throw error;
    }
  }, []);

  // Manual chart deactivation
  const deactivateChart = useCallback(() => {
    setActiveChartId(null);
    setIsChartActive(false);
    
    if (reconnectionTimer.current) {
      clearTimeout(reconnectionTimer.current);
      reconnectionTimer.current = null;
    }
  }, []);

  // Refresh chart selection (especially useful for Excel Online)
  const refreshChartSelection = useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        // Try to get currently selected chart
        const selection = context.workbook.getSelectedRange();
        selection.load(['address']);
        
        // Check if any chart is currently active
        const charts = context.workbook.charts;
        charts.load(['items']);
        
        await context.sync();
        
        // Look for an active chart
        let foundActiveChart = false;
        
        for (const chart of charts.items) {
          chart.load(['id', 'name']);
          await context.sync();
          
          try {
            // Try to determine if this chart is currently selected/active
            // This is a workaround since there's no direct API to check selection
            const chartRange = chart.getImage();
            await context.sync();
            
            // If we can access the chart without error, it might be active
            if (activeChartId !== chart.id) {
              setActiveChartId(chart.id);
              setIsChartActive(true);
              foundActiveChart = true;
              break;
            }
          } catch (error) {
            // Chart is not active/accessible
          }
        }
        
        if (!foundActiveChart && isChartActive) {
          setActiveChartId(null);
          setIsChartActive(false);
        }
      });
    } catch (error) {
      console.error('Chart selection refresh failed:', error);
    }
  }, [activeChartId, isChartActive]);

  // Initialize events when component mounts
  useEffect(() => {
    const initializeEvents = async () => {
      await registerChartEvents();
      
      // For Excel Online, set up periodic refresh
      if (isOnlineRef.current) {
        const refreshInterval = setInterval(() => {
          refreshChartSelection();
        }, 2000); // Check every 2 seconds
        
        return () => clearInterval(refreshInterval);
      }
    };

    // Register worksheet events to detect new charts
    const registerWorksheetEvents = async () => {
      try {
        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          
          // Register onActivated for worksheets to re-register chart events
          const worksheetHandler = worksheets.onActivated.add(async () => {
            // Small delay to ensure worksheet is fully loaded
            setTimeout(() => {
              registerChartEvents();
            }, 300);
          });
          
          await context.sync();
          
          // Store worksheet handler
          eventHandlers.current.set('worksheet-activated', worksheetHandler);
        });
      } catch (error) {
        console.error('Worksheet event registration failed:', error);
      }
    };

    initializeEvents();
    registerWorksheetEvents();

    // Cleanup on unmount
    return () => {
      eventHandlers.current.forEach((handler) => {
        try {
          handler.remove();
        } catch (error) {
          console.warn('Handler cleanup failed:', error);
        }
      });
      eventHandlers.current.clear();
      
      if (reconnectionTimer.current) {
        clearTimeout(reconnectionTimer.current);
      }
    };
  }, [registerChartEvents, refreshChartSelection]);

  const contextValue: ChartContextValue = {
    activeChartId,
    isChartActive,
    activateChart,
    deactivateChart,
    refreshChartSelection,
  };

  return (
    <ChartContext.Provider value={contextValue}>
      {children}
    </ChartContext.Provider>
  );
};
```
