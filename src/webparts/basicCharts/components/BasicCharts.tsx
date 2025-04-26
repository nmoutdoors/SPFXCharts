import * as React from 'react';
import { useState, useEffect } from 'react';
import type { IBasicChartsProps } from './IBasicChartsProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { Pie, Bar, Line, Radar, Scatter, Doughnut, PolarArea } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  LineElement,
  PointElement,
  RadialLinearScale,
  TimeScale
} from 'chart.js';

// Register Chart.js components
ChartJS.register(
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  LineElement,
  PointElement,
  RadialLinearScale,
  TimeScale  // Important for scatter plot with time axis
);

interface IGrade {
  ID: number;
  Title: string;
  Grade: number;
  Date: string;
}

type TabType = 'pie' | 'bar' | 'line' | 'radar' | 'scatter' | 'doughnut' | 'polar';

export default function BasicCharts(props: IBasicChartsProps): React.ReactElement<IBasicChartsProps> {
  const [grades, setGrades] = useState<IGrade[]>([]);
  const [error, setError] = useState<string>('');
  const [activeTab, setActiveTab] = useState<TabType>('pie');

  const loadGrades = async (): Promise<void> => {
    try {
      const listUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Grades-N-Such')/items?$select=ID,Title,Grade,Date`;
      
      const response = await props.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '3.0'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      setGrades(data.d.results);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      console.error('Error in loadGrades:', errorMessage);
      setError(errorMessage);
    }
  };

  useEffect(() => {
    loadGrades().catch(error => {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      console.error('Error in useEffect:', errorMessage);
      setError(errorMessage);
    });
  }, []);

  // Common colors for charts
  const colors = [
    '#FF6384',
    '#36A2EB',
    '#FFCE56',
    '#4BC0C0',
    '#9966FF',
    '#FF9F40',
    '#FF6384',
    '#36A2EB',
    '#FFCE56',
    '#4BC0C0'
  ];

  // Add these new color arrays specifically for the polar chart
  const polarColors = [
    '#4BC0C0', // teal
    '#FF9F40', // orange
    '#36A2EB', // blue
    '#FF6384', // pink
    '#9966FF', // purple
    '#FFCE56', // yellow
    '#4BCFFF', // light blue
    '#FF99CC', // light pink
    '#99FF99', // light green
    '#FFB366'  // light orange
  ];

  const polarBorderColors = [
    '#3AA0A0', // darker teal
    '#E68A30', // darker orange
    '#2B82C9', // darker blue
    '#E84D6D', // darker pink
    '#7F47FF', // darker purple
    '#E6B840', // darker yellow
    '#3AAFDD', // darker light blue
    '#FF77AA', // darker light pink
    '#77DD77', // darker light green
    '#FF9940'  // darker light orange
  ];

  // Common options for all charts
  const commonOptions = {
    responsive: true,
    maintainAspectRatio: true,
    plugins: {
      legend: {
        position: 'bottom' as const
      }
    }
  };

  // Chart data configurations
  const pieChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      data: grades.map(grade => grade.Grade),
      backgroundColor: colors,
      hoverBackgroundColor: colors
    }]
  };

  const barChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      label: 'Grade',
      data: grades.map(grade => grade.Grade),
      backgroundColor: colors[1],
      borderColor: colors[1],
      borderWidth: 1
    }]
  };

  const lineChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      label: 'Grade',
      data: grades.map(grade => grade.Grade),
      fill: false,
      borderColor: colors[2],
      tension: 0.1
    }]
  };

  const radarChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      label: 'Grade',
      data: grades.map(grade => grade.Grade),
      backgroundColor: `${colors[3]}44`,
      borderColor: colors[3],
      borderWidth: 2
    }]
  };

  const scatterChartData = {
    datasets: [{
      label: 'Grades Over Time',
      data: grades.map(grade => ({
        x: new Date(grade.Date).valueOf(), // Convert to timestamp
        y: grade.Grade
      })),
      backgroundColor: colors[4],
      pointRadius: 6,
      pointHoverRadius: 8
    }]
  };

  const scatterChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { 
        display: true, 
        text: 'Grades by Date' 
      }
    },
    scales: {
      x: {
        type: 'linear' as const,
        position: 'bottom' as const,
        title: {
          display: true,
          text: 'Date'
        },
        ticks: {
          callback: function(value: number) {
            return new Date(value).toLocaleDateString();
          }
        }
      },
      y: {
        beginAtZero: true,
        max: 100,
        ticks: {
          stepSize: 10
        },
        title: {
          display: true,
          text: 'Grade'
        }
      }
    },
    responsive: true,
    maintainAspectRatio: true,
    aspectRatio: 2
  };

  const doughnutChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      data: grades.map(grade => grade.Grade),
      backgroundColor: colors,
      hoverBackgroundColor: colors
    }]
  };

  const polarChartData = {
    labels: grades.map(grade => grade.Title),
    datasets: [{
      data: grades.map(grade => grade.Grade),
      backgroundColor: polarColors.map(color => `${color}88`), // Adding transparency
      borderColor: polarBorderColors,
      borderWidth: 2
    }]
  };

  // Chart-specific options
  const pieChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Distribution (Pie)' }
    }
  };

  const barChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Distribution (Bar)' }
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 100
      }
    }
  };

  const lineChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Progression' }
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 100
      }
    }
  };

  const radarChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Radar Analysis' }
    },
    scales: {
      r: {
        beginAtZero: true,
        max: 100
      }
    }
  };

  const doughnutChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Distribution (Doughnut)' }
    }
  };

  const polarChartOptions = {
    ...commonOptions,
    plugins: {
      ...commonOptions.plugins,
      title: { display: true, text: 'Grade Distribution (Polar)' }
    }
  };

  const tabStyle = {
    padding: '8px 16px',
    margin: '0 2px',
    border: 'none',
    borderRadius: '4px 4px 0 0',
    cursor: 'pointer',
    fontSize: '13px',
    fontWeight: 'bold' as const,
  };

  const activeTabStyle = {
    ...tabStyle,
    backgroundColor: '#f0f0f0',
    borderBottom: '2px solid #0078d4',
  };

  const inactiveTabStyle = {
    ...tabStyle,
    backgroundColor: '#e0e0e0',
    borderBottom: '2px solid transparent',
  };

  return (
    <div>
      <h2>SPFX and Chart.js</h2>
      {error && <div style={{ color: 'red', marginBottom: '1em' }}>Error: {error}</div>}
      {grades.length === 0 && !error && <div>Loading...</div>}
      
      {/* Tabs Navigation */}
      <div style={{ marginBottom: '20px', borderBottom: '1px solid #ccc', display: 'flex', flexWrap: 'wrap' }}>
        {[
          { id: 'pie', label: 'Pie Chart' },
          { id: 'bar', label: 'Bar Chart' },
          { id: 'line', label: 'Line Chart' },
          { id: 'radar', label: 'Radar Chart' },
          { id: 'scatter', label: 'Scatter Plot' },
          { id: 'doughnut', label: 'Doughnut Chart' },
          { id: 'polar', label: 'Polar Area' }
        ].map(tab => (
          <button
            key={tab.id}
            style={activeTab === tab.id ? activeTabStyle : inactiveTabStyle}
            onClick={() => setActiveTab(tab.id as TabType)}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* Chart Container */}
      <div style={{ minHeight: '400px', marginBottom: '2em' }}>
        <div style={{ maxWidth: '800px', margin: '0 auto' }}>
          {activeTab === 'pie' && <Pie data={pieChartData} options={pieChartOptions} />}
          {activeTab === 'bar' && <Bar data={barChartData} options={barChartOptions} />}
          {activeTab === 'line' && <Line data={lineChartData} options={lineChartOptions} />}
          {activeTab === 'radar' && <Radar data={radarChartData} options={radarChartOptions} />}
          {activeTab === 'scatter' && <Scatter data={scatterChartData} options={scatterChartOptions} />}
          {activeTab === 'doughnut' && <Doughnut data={doughnutChartData} options={doughnutChartOptions} />}
          {activeTab === 'polar' && <PolarArea data={polarChartData} options={polarChartOptions} />}
        </div>
      </div>

      {/* Data Table */}
      <div style={{ border: '1px solid #ccc', padding: '10px' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              <th style={{ border: '1px solid #ddd', padding: '8px', textAlign: 'left' }}>ID</th>
              <th style={{ border: '1px solid #ddd', padding: '8px', textAlign: 'left' }}>Title</th>
              <th style={{ border: '1px solid #ddd', padding: '8px', textAlign: 'left' }}>Grade</th>
              <th style={{ border: '1px solid #ddd', padding: '8px', textAlign: 'left' }}>Date</th>
            </tr>
          </thead>
          <tbody>
            {grades.map(grade => (
              <tr key={grade.ID}>
                <td style={{ border: '1px solid #ddd', padding: '8px' }}>{grade.ID}</td>
                <td style={{ border: '1px solid #ddd', padding: '8px' }}>{grade.Title}</td>
                <td style={{ border: '1px solid #ddd', padding: '8px' }}>{grade.Grade}</td>
                <td style={{ border: '1px solid #ddd', padding: '8px' }}>
                  {new Date(grade.Date).toLocaleDateString()}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: '1em' }}>
        Total records: {grades.length}
      </div>
    </div>
  );
}








